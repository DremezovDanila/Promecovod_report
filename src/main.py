
# Program realises PDF reports for crucial parameters of stations.

from pyModbusTCP.client import ModbusClient
from pyModbusTCP import utils
from fpdf import FPDF
import time
import datetime
import os
from smb.SMBConnection import SMBConnection
import smbclient
import subprocess
from typing import Any
from threading import Thread
from calendar import monthrange

""" 
    Lines in ini.txt file start from 0,
    lines in internal table (list of lists) start from 0,
    lines in PDF report start from 1.
"""

# Create class PDF based on original PDF class from FPDF2 library.
class Pdf_2(FPDF):
    # Method creating pdf document with table.
    def create_report(self, datetime: datetime, margins: tuple, table_data: list[list[Any, ...], ...], num_params: int):
        # Format pdf document.
        self.set_margins(margins[0], margins[1], margins[2])
        # Create the first page.
        self.add_page()
        # Download the font supporting Unicode and set it. *We have to add a font before setting it.
        self.add_font("font_1", "", r"..\etc\font\ARIALUNI.ttf")
        # Upper line with organization name on the right and current date on the left.
        self.set_font("font_1", size=14)
        self.cell(w=(self.w - (margins[0] + abs(margins[2]))) / 2, h=self.font_size, txt="{:%Y.%m.%d}".format(datetime), align="L")
        self.cell(w=(self.w - (margins[0] + abs(margins[2]))) / 2, h=self.font_size, txt='ООО "ПромЭкоВод"', align="R", new_x="LMARGIN",
                  new_y="NEXT")
        # Padding.
        self.cell(w=self.w, h=10, new_x="LMARGIN", new_y="NEXT")
        # Site name.
        self.set_font("font_1", size=20)
        self.cell(self.w - 20, self.font_size * 1.5, txt=table_data[0][1], align="C", new_x="LEFT",
                  new_y="NEXT")
        # Table title.
        self.set_font("font_1", size=14)
        self.cell(self.w - 20, self.font_size * 2, txt=f"Отчет по суточным расходам воды и электроэнергии",
                  align="C", new_x="LEFT", new_y="NEXT")
        self.cell(w=self.w, h=2, new_x="LMARGIN", new_y="NEXT")

        # Draw a table.

        # Declare variables.
        head_height = 5
        row_height = 6.5

        # Get table_row_mask list from table_data.
        table_line_mask = table_data[2]

        # Find widths of columns in table using specified font.
        self.set_font("font_1", size=14)
        table_column_width = []
        for item in table_line_mask:
            table_column_width.append(self.get_string_width(item))

        # Get table_head list from table_data.
        table_head = table_data[1]
        # Replace character \n (\ and n) to \u000A (\n - Line Feed) in table_head variable.
        table_head_f = []
        for item in table_head:
            table_head_f.append(item.replace(r"\n", "\n"))

        # Draw a table head.
        self.set_font("font_1", size=12)
        # Number of parameters plus 2 general parameters.
        for column_num in range(num_params + 2):
            self.multi_cell(w=table_column_width[column_num] * 1.1, h=head_height, txt=table_head_f[column_num],
                            border=1, align="C", new_x="RIGHT",new_y="TOP")
        # Line break (Find out how many rows table head takes)
        # if table_head_f[0].count('\n') > 0:
        #     for i in range(table_head_f[0].count('\n')):
        #         self.multi_cell(w=0, h=head_height, txt="", border=0, align="C", new_x="RIGHT",
        #                         new_y="NEXT")
        # else:
        #     self.multi_cell(w=0, h=head_height, txt="", border=0, align="C", new_x="RIGHT",
        #                     new_y="NEXT")

        # Line break (Table head takes 4 lines).
        for i in range(4):
            self.multi_cell(w=0, h=head_height, txt="", border=0, align="C", new_x="RIGHT",
                            new_y="NEXT")
        # Set the carriage to left margin. Left y coordinate the same.
        self.set_x(self.l_margin)

        # Draw the table (line starts from 3rd line in internal table (data list)).
        self.set_font("font_1", size=14)
        for line_num in range(3, len(table_data)):
            # for column_num in range(len(table_data[line_num])):
            # 1 column - line number, 2 column - date and number of params (columns in line).
            for column_num in range(num_params + 2):
                # Increase Line number in PDF to 1 relating to internal table.
                # if column_num == 0:
                #     try:
                #         table_value = str(int(table_data[line_num][column_num]) + 1)
                #     except:
                #         table_value = table_data[line_num][column_num]
                # else:
                #     table_value = table_data[line_num][column_num]

                self.cell(w=table_column_width[column_num] * 1.1, h=row_height,
                          txt=f"{table_data[line_num][column_num]}", border=1, align="C", new_x="RIGHT",
                          new_y="TOP")
            # Break table row
            self.cell(w=0, h=row_height, txt="", border=0, align="C", new_x="RIGHT",
                      new_y="NEXT")
            self.set_x(self.l_margin)

    # Override footer according to requires.
    def footer(self):
        self.set_y(-15)
        self.set_x(10)
        self.set_font("font_1", "", 12)
        self.cell(w=self.w - 20, h=10, txt='Стр. %s из ' % self.page_no() + '{nb}', align='C')


# This function fill in internal table with meanings from ini.txt file.
# Inputs: name of site, path to ini file, internal table like empty list.
# Outputs: active line in ini.txt (and in internal table too).
# Isn't used currently.
def initialize_table(obj_params: tuple, ini_txt_path: str, table_data: list, pdf_table_head: tuple, pdf_table_mask: tuple) -> tuple:
    # Check the existence of ini.txt file. If exists, then read it into table data.
    if os.path.isfile(ini_txt_path):
        output_message = f"{obj_params[1]}. The initial txt file has been existing."
        with open(ini_txt_path, "r", encoding="UTF-8") as ini_txt_file:
            ini_txt_file_list_lines = ini_txt_file.readlines()
            # Count lines in the ini TXT file. The last one is active.
            ini_txt_file_active_line = len(ini_txt_file_list_lines)
            # if len(ini_txt_file_list_lines) > 3:
            for line in ini_txt_file_list_lines:
                ini_txt_line = line.split(";")
                ini_txt_line.pop(-1)
                table_data.append(ini_txt_line)
    # If ini.txt file doesn't exist, then create it empty.
    else:
        output_message = f"{obj_params[1]}. The initial txt file has been created."
        with open(ini_txt_path, "w", encoding="UTF-8") as ini_txt_file:
            # Write parameters into ini.txt.
            for item in obj_params:
                ini_txt_file.write(item + ";")
            ini_txt_file.write("\n")
            # Write head.
            for item in pdf_table_head:
                ini_txt_file.write(item + ";")
            ini_txt_file.write("\n")
            # Write mask.
            for item in pdf_table_mask:
                ini_txt_file.write(item + ";")
            ini_txt_file.write("\n")
            ini_txt_file_active_line = 3
        # Read parameters into internal table.
        table_data.append(obj_params)
        table_data.append(pdf_table_head)
        table_data.append(pdf_table_mask)
    # Return number of current active line in ini.txt.
    return ini_txt_file_active_line, output_message


# Read data from ini.txt file into internal table.
def read_ini_txt(ini_txt_path: str, data: list) -> int:
    # Check the existence of ini.txt file. If exists, then read it into table data.
    if os.path.isfile(ini_txt_path):
        with open(ini_txt_path, "r", encoding="UTF-8") as ini_txt_file:
            ini_txt_file_list_lines = ini_txt_file.readlines()
            # Count lines in the ini TXT file. The last one is active.
            ini_txt_file_active_line = len(ini_txt_file_list_lines) + 1
            # If a file isn't empty.
            # if len(ini_txt_file_list_lines) > 0:
            for line in ini_txt_file_list_lines:
                ini_txt_line = line.split(";")
                # Delete "\n" special symbol at the end of line.
                ini_txt_line.pop(-1)
                data.append(ini_txt_line)
    else:
        ini_txt_file_active_line = 0
    return ini_txt_file_active_line


# If ini.txt file doesn't exist, then create it filling lines for previous days of month by "0.0".
def create_ini_txt(ini_txt_path: str, object: list, day_prev_num: int) -> str:
    with open(ini_txt_path, "w", encoding="UTF-8") as ini_txt_file:
        # Write parameters into ini.txt.
        obj_params = object[3][0]
        for item in obj_params:
            ini_txt_file.write(item + ";")
        ini_txt_file.write("\n")
        # Write head.
        pdf_table_head = object[3][1]
        for item in pdf_table_head:
            ini_txt_file.write(item + ";")
        ini_txt_file.write("\n")
        # Write mask.
        pdf_table_mask = object[3][2]
        for item in pdf_table_mask:
            ini_txt_file.write(item + ";")
        ini_txt_file.write("\n")
        # Fill lines of previous days in month. When 1st day in month report is being printed, day_prev_num = 0 and
        # no lines have to be filled in.
        for day in range(day_prev_num):
            ini_txt_file.write(f"{day + 1};")
            ini_txt_file.write(f"{datetime.datetime.today().year}.{datetime.datetime.today().month:02}."
                               f"{day + 1:02};")
            for item in range(5):
                ini_txt_file.write("0.0;")
            ini_txt_file.write("\n")
        logs_message = f"{object[0][0]}. The initial txt file has been created."
    # Return number of current active line in ini.txt.
    return logs_message


# Check a directory for saving pdf report.
def check_pdf_dir(obj_params: tuple):
    if os.path.isdir(rf"c:\Отчетность по работе станций\{obj_params[1]}"):
        output_message = rf"The directory 'x:\Отчетность по работе станций\{obj_params[1]}' exists."
        directory_exist = True
    else:
        try:
            os.makedirs(rf"c:\Отчетность по работе станций\{obj_params[1]}")
            output_message = rf"Directory 'x:\Отчетность по работе станций\{obj_params[1]}' was created."
            directory_exist = True
        except:
            output_message = rf"Directory 'x:\Отчетность по работе станций\{obj_params[1]}' error while being created."
            directory_exist = False
    return directory_exist, output_message


# Get active line in ini.txt.
def get_active_line_txt(ini_txt_path: str):
    with open(ini_txt_path, "r", encoding="UTF-8") as ini_txt:
        ini_txt_num_lines = len(ini_txt.readlines())
        # First 3 lines takes common information about an object.
        if ini_txt_num_lines == 3:
            ini_txt_active_line = 1
        else:
            ini_txt_active_line = ini_txt_num_lines - 3 + 1
    return ini_txt_active_line


# Using commands in console to operate program.
def operate_program():
    global finish_main_process
    global restore_start
    global restore_ini_txt_name
    global restore_ini_txt_date
    global print_start
    global print_ini_txt_name
    global print_ini_txt_date
    while not finish_main_process:
        print("Enter your command:\n")
        command = input()
        if command == "finish" or command == "Finish" or command == "FINISH":
            finish_main_process = True
            command = ""
        elif command == "restore" or command == "Restore" or command == "RESTORE":
            restore_ini_txt_name = str(input("Name of object (ex.:vzu_borodinsky): "))
            restore_ini_txt_date = str(input("Year and month (ex.:2023_05): "))
            restore_start = True
            command = ""
        elif command == "print" or command == "Print" or command == "PRINT":
            print_ini_txt_name = str(input("Name of object (ex.:vzu_borodinsky): "))
            print_ini_txt_date = str(input("Year and month (ex.:2023_05): "))
            print_start = True
            command = ""


# Global variables.
restore_start = False                           # Command to start pdf file restoring.
print_start = False                             # Command to start pdf file printing
finish_main_process = False                     # Command to finish program.
# Parameters for pdf formatting.
pdf_left_margin = 10
pdf_top_margin = 10
pdf_right_margin = -10
pdf_margins = (pdf_left_margin, pdf_top_margin, pdf_right_margin)
# Initialize current date and time.
cur_datetime = datetime.datetime.now()
cur_date = cur_datetime.date()
cur_time = cur_datetime.time()
cur_time_hour = cur_time.hour
cur_time_min = cur_time.minute
cur_time_sec = cur_time.second
cur_date_day = datetime.datetime.today().day
cur_date_month = datetime.datetime.today().month
cur_date_year = datetime.datetime.today().year
report_cur_datetime = cur_datetime
# Logs
logs_num_line = 0
logs_separator = " " * 4

# VZU Borodinsky (vboro).

vboro_num_params = 5                            # Number of parameters of object.
# Information about a project for ini.txt.
vboro_obj_params = ("VZU Borodinsky",
                    "ВЗУ Бородинский",
                    r"x:\Отчетность по работе станций\ВЗУ Бородинский",
                    r"..\reports\vzu_borodinsky",
                    "vzu_borodinsky",
                    "",
                    ""
                    )
# The header of pdf report.
vboro_pdf_table_head = (r"\n№\n\n\n",
                        r"\nДата/время\n\n\n",
                        r"\nВоды за сутки, м3\n\n",
                        r"\nВоды всего, м3\n\n",
                        r"\nЭнергии за сутки, кВт\n\n",
                        r"\nЭнергии всего, кВт\n\n",
                        r"\nЭнерг. на куб, кВт/м3\n\n"
                        )
# Set the mask for pdf table columns.
vboro_pdf_table_mask = ("00000", "0000.00.00  00:00", "0000000.0", "0000000.0", "0000000.0",
                        "0000000.0", "0000000.00")
# List of ini.txt parameters
vboro_ini_txt_params = [vboro_obj_params, vboro_pdf_table_head, vboro_pdf_table_mask]
# MBTCP parameters
vboro_slave_address = '192.168.239.50'
vboro_port = 503
vboro_unit_id = 1
vboro_timeout = 15.0
vboro_mbtcp_params = (vboro_slave_address, vboro_port, vboro_unit_id, vboro_timeout)
# List of vboro object parameters.
vboro_common_params = [vboro_obj_params, vboro_num_params, vboro_mbtcp_params, vboro_ini_txt_params]    # Common parameters for object.


# VZU Teploe (vtepl).

vtepl_num_params = 3                            # Number of parameters of object.
# Information about project for ini.txt.
vtepl_obj_params = ("VZU Teploe",
                    "ВЗУ Теплое",
                    r"c:\Отчетность по работе станций\ВЗУ Теплое",
                    r"..\reports\vzu_teploe",
                    "vzu_teploe",
                    "",
                    ""
                    )
# The header of pdf report.
vtepl_pdf_table_head = (r"\n№\n\n\n",
                        r"\nДата/время\n\n\n",
                        r"\nРасход на входе 1, м3\n\n",
                        r"\nРасход на входе 2, м3\n\n",
                        r"\nРасход на выходе, м3\n\n",
                        "",
                        ""
                        )
# Set the mask for pdf table columns.
vtepl_pdf_table_mask = ("00000", "0000.00.00  00:00", "0000000000000.0", "0000000000000.0", "0000000000000.0", "", "")
# Parameters of ini.txt.
vtepl_ini_txt_params = [vtepl_obj_params, vtepl_pdf_table_head, vtepl_pdf_table_mask]
# MBTCP parameters.
vtepl_slave_address = '192.168.239.50'
vtepl_port = 503
vtepl_unit_id = 1
vtepl_timeout = 15.0
vtepl_mbtcp_params = (vtepl_slave_address, vtepl_port, vtepl_unit_id, vtepl_timeout)
# List of vtepl object parameters.
vtepl_common_params = [vtepl_obj_params, vtepl_num_params, vtepl_mbtcp_params, vtepl_ini_txt_params]    # Common parameters for object.


# KOS Makarovo  (kmkrv).

kmkrv_num_params = 3                            # Number of parameters of object.
# Information about project for ini.txt.
kmkrv_obj_params = ("KOS Makarovo",
                    "КОС Макарово",
                    r"c:\Отчетность по работе станций\КОС Макарово",
                    r"..\reports\kos_makarovo",
                    "kos_makarovo",
                    "",
                    ""
                    )
# The header of pdf report.
kmkrv_pdf_table_head = (r"\n№\n\n\n",
                        r"\nДата/время\n\n\n",
                        r"\nРасход на входе 1, м3\n\n",
                        r"\nРасход на входе 2, м3\n\n",
                        r"\nРасход на выходе, м3\n\n",
                        "",
                        ""
                        )
# Set the mask for pdf table columns.
kmkrv_pdf_table_mask = ("00000", "0000.00.00  00:00", "0000000000000.0", "0000000000000.0", "0000000000000.0", "", "")
# Parameters of ini.txt.
kmkrv_ini_txt_params = [kmkrv_obj_params, kmkrv_pdf_table_head, kmkrv_pdf_table_mask]
# MBTCP parameters.
kmkrv_slave_address = '192.168.239.50'
kmkrv_port = 503
kmkrv_unit_id = 1
kmkrv_timeout = 15.0
kmkrv_mbtcp_params = (kmkrv_slave_address, kmkrv_port, kmkrv_unit_id, kmkrv_timeout)
# List of kmkrv object parameters.
kmkrv_common_params = [kmkrv_obj_params, kmkrv_num_params, kmkrv_mbtcp_params, kmkrv_ini_txt_params]    # Common parameters for object.


# List of objects.
objects_com_params = [vboro_common_params, vtepl_common_params, kmkrv_common_params]

# Check existence of ini.txt file.
for object in objects_com_params:
    # Get path to object ini.txt.
    object_ini_txt_path = rf"{object[0][3]}/{object[0][4]}_ini_{cur_date_year}_{cur_date_month:02}.txt"
    # If exists, then check number of lines.
    if os.path.isfile(object_ini_txt_path):
        with open(f"{object_ini_txt_path}", "r", encoding="UTF-8") as object_ini_txt:
            ini_txt_file_list_lines = object_ini_txt.readlines()
            # If file has 3 or more lines it's ok.
            if len(ini_txt_file_list_lines) > 2:
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(
                        f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}{object[0]}_ini.txt exists\n")
            else:
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(
                        f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}{object[0]}_ini.txt is off\n")
    # Else, renew file.
    else:
        create_ini_txt(object=object, ini_txt_path=object_ini_txt_path, day_prev_num=cur_date_day - 1)

# Start another thread to process operating commands.
th = Thread(target=operate_program, args=())
th.start()

# Main loop.
if __name__ == "__main__":
    while not finish_main_process:
        # Get current date and time.
        cur_datetime = datetime.datetime.now()
        cur_date = cur_datetime.date()
        cur_time = cur_datetime.time()
        cur_time_hour = cur_time.hour
        cur_time_min = cur_time.minute
        cur_time_sec = cur_time.second
        cur_date_day = datetime.datetime.today().day
        cur_date_month = datetime.datetime.today().month
        cur_date_year = datetime.datetime.today().year

        # Start polls once a day at specified time (next day after report get formed in PLC). PLC draw report at
        # 23:59:59 and client polls start several minutes later.
        if cur_time_hour == 0 and cur_time_min == 5 and cur_time_sec == 0:

            # Get date of previous day (report).
            if cur_date_day == 1 and cur_date_month == 1:
                report_last_year = cur_date_year - 1
                report_last_month = 1
                report_last_day = 31
            elif cur_date_day == 1 and cur_date_month != 1:
                report_last_year = cur_date_year
                report_last_month = cur_date_month - 1
                # Last day of previous month.
                report_last_day = monthrange(int(report_last_year), int(report_last_month))[1]
            else: # elsif cur_date_day != 1:
                report_last_year = cur_date_year
                report_last_month = cur_date_month
                report_last_day = cur_date_day - 1
            # Last day datetime (when last report got formed).
            report_last_datetime = report_cur_datetime
            # New day datetime (when next report get formed).
            report_cur_datetime = datetime.datetime.now()

            # To interact with network drive we have to add it to the local directory by mapping it to the letter (X:).
            # Check whether drive X has already created. If not, then create network drive X.
            if not os.path.isdir("x:"):
                # The command sets utf-8 (65001 code page) for cmd.exe.
                command_1 = "chcp 65001"
                cmd_command_1 = subprocess.run(command_1, shell=True, capture_output=True)
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               f"Change encoding: {cmd_command_1.stdout}\n")
                # The command maps network share to drive letter X on the local system by SMB with authorization.
                command_2 = r'net use x: "\\192.168.1.2\ics" ddd2232 /user:dremezov /persistent:yes'
                try:
                    cmd_command_2 = subprocess.run(command_2, shell=True, capture_output=True)
                    with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"Create network drive X: {cmd_command_2.stdout}\n")
                # Process exception while sys command performing in subprocess.
                except subprocess.CalledProcessError:
                    with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"Error of creating drive X.\n")
            else:
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               f"Drive X exists.\n")

            # ModbusTCP polling.
            for object in objects_com_params:
                # Set path to object's ini.txt file.
                object_ini_txt_path = f"{object[3][0][3]}/{object[3][0][3]}_ini_{report_last_year}_{report_last_month:02}.txt"
                # Check existence of ini.txt file. If exists then move on, else create new ini.txt file.
                if not os.path.isfile(object_ini_txt_path):
                    create_ini_txt(object=object, ini_txt_path=object_ini_txt_path, day_prev_num=report_last_day - 1)
                # Parametrise MBTCP connection.
                object_mbtcp_client = ModbusClient(host=object[2][0],
                                   port=int(object[2][1]),
                                   unit_id=int(object[2][2]),
                                   timeout=float(object[2][3]),
                                   auto_open=True,
                                   auto_close=True)
                # List of polled parameters for object.
                object_parameters = []
                # Add line number and date to "object_parameters" variable.
                object_ini_txt_active_line = get_active_line_txt(object_ini_txt_path)
                object_parameters.append(object_ini_txt_active_line)
                object_parameters.append(f"{report_last_year}.{report_last_month:02}.{report_last_day:02}")
                # Read parameters for every object one by one.
                for poll in range(int(object[1])):
                    # Write year and month of report and parameter number (returns True or False).
                    object_poll_w_status = object_mbtcp_client.write_multiple_registers(0, [report_last_year,
                                                                                            report_last_month,
                                                                                            poll + 1])
                    # Read parameter.
                    object_poll_r = object_mbtcp_client.read_input_registers((report_last_day - 1) * 2, 2)
                    # If poll fulfilled successfully, save parameter value.
                    if object_poll_w_status:
                        # Conversion list of 2int16 to list of long32.
                        object_poll_r_long_32 = utils.word_list_to_long(object_poll_r, big_endian=True)
                        # Conversion item of list of long32 to real.
                        object_poll_r_real = utils.decode_ieee(object_poll_r_long_32[0])
                        # Get meaning and format it, then add to list of parameters.
                        object_parameters.append(format(object_poll_r_real, '.1f'))
                    else:
                        # Else write value as 0.0.
                        object_parameters.append(format(0.0, '.1f'))
                        with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                            # object[3][0][0] - object name.
                            logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                       f"{object[3][0][0]}. MBTCP error: {object_mbtcp_client.last_error}\n")

                # Check if the line (list) filled fully. It must have 7 rows.
                object_parameters_len = len(object_parameters)
                if object_parameters_len < 7:
                    for i in range(7 - object_parameters_len):
                        object_parameters.append(format(0.0, '.1f'))
                # Write new row of data into ini.txt file.
                with open(object_ini_txt_path, "a", encoding="UTF-8") as object_ini_txt:
                    for elem in object_parameters:
                        object_ini_txt.write(f"{elem};")
                    object_ini_txt.write("\n")

            # PDF document printing.

            for object in objects_com_params:
                report_pdf = Pdf_2()
                # Set path to object's ini.txt file.
                object_ini_txt_path = f"{object[3][0][3]}/{object[3][0][3]}_ini_{report_last_year}_{report_last_month:02}.txt"
                # print(object_ini_txt_path)
                # Read data from ini.txt and write it into list2.
                object_data_list2 = []
                read_ini_txt(object_ini_txt_path, object_data_list2)
                # print(object_data_list2)
                # Fill in pdf.
                report_pdf.create_report(datetime=report_last_datetime, margins=pdf_margins,
                                         table_data=object_data_list2, num_params=object[1])
                # Check existence of pdf report directory.
                check_pdf_dir_return = check_pdf_dir(object[3][0])
                pdf_dir_exist = check_pdf_dir_return[0]
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               f"{check_pdf_dir_return[1]}\n")
                if pdf_dir_exist:
                    report_pdf.output(
                        rf"c:\Отчетность по работе станций\{object[3][0][1]}\Отчет по {object[3][0][1]} от {report_last_year}_{report_last_month:02}.pdf")

        # Print a report for specified time if it was deleted.
        if print_start:
            print_start = False
            # Checking input restoring date for adequacy.
            try:
                print_ini_txt_date_list = print_ini_txt_date.split("_")
                print_ini_txt_date_year = print_ini_txt_date_list[0]
                print_ini_txt_date_year_int = int(print_ini_txt_date_year)
                print_ini_txt_date_month = print_ini_txt_date_list[1]
                print_ini_txt_date_month_int = int(print_ini_txt_date_month)
                if len(print_ini_txt_date_list) == 2 and \
                        len(print_ini_txt_date_year) == 4 and \
                        (len(print_ini_txt_date_month) == 2 or len(print_ini_txt_date_month) == 1) and \
                        2022 < int(print_ini_txt_date_year) < 2040 and \
                        (1 <= int(print_ini_txt_date_month) <= 12):
                    # Path to ini.txt file.
                    print_path = rf"..\reports\{print_ini_txt_name}\{print_ini_txt_name}_ini_{print_ini_txt_date_year_int}_" \
                                 rf"{print_ini_txt_date_month_int:02}.txt"
                    if os.path.isfile(print_path):
                        # Get parameters number of object.
                        print_num_params = 0
                        for object in objects_com_params:
                            if object[3][0][4] == print_ini_txt_name:
                                print_num_params = object[1]
                        with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                            logs.write(rf"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                       rf"{print_ini_txt_name} report for {print_ini_txt_date} printing started." + "\n")
                        # Create new data list.
                        print_data = []
                        # Read ini.txt file needed to restore report.
                        read_ini_txt(ini_txt_path=print_path, data=print_data)
                        # Get general params from internal table.
                        print_object_params = print_data[0]
                        print_pdf_table_head = print_data[1]
                        print_pdf_table_mask = print_data[2]
                        # Create new pdf for restore.
                        pdf_print = Pdf_2()
                        # Create report.
                        pdf_print.create_report(datetime=cur_datetime, margins=pdf_margins, table_data=print_data,
                                                num_params=print_num_params)
                        # Check existence of pdf report directory.
                        check_print_pdf_dir_return = check_pdf_dir(print_object_params)
                        print_pdf_dir_exist = check_print_pdf_dir_return[0]
                        with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                            logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                       f"{check_print_pdf_dir_return[1]}\n")
                        if print_pdf_dir_exist:
                            pdf_print.output(
                                rf"c:\Отчетность по работе станций\{print_object_params[1]}\Отчет по {print_object_params[1]} "
                                rf"от {print_ini_txt_date_year_int}_{print_ini_txt_date_month_int:02}_восстановленный.pdf")
                    else:
                        with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                            logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                       f"Printing. File in {print_path} doesn't exist.\n")
            except:
                print_check_ok = False
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(rf"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               rf"Printing. Exception while checking input data ({print_ini_txt_name}_{print_ini_txt_date_year_int})." + "\n")


        # Restore ini.txt of specified date from PLC.
        if restore_start:
            restore_start = False
            # Checking input restoring date for adequacy.
            try:
                restore_ini_txt_date_list = restore_ini_txt_date.split("_")
                restore_ini_txt_date_year = restore_ini_txt_date_list[0]
                restore_ini_txt_date_year_int = int(restore_ini_txt_date_year)
                restore_ini_txt_date_month = restore_ini_txt_date_list[1]
                restore_ini_txt_date_month_int = int(restore_ini_txt_date_month)
                if len(restore_ini_txt_date_list) == 2 and \
                    len(restore_ini_txt_date_year) == 4 and \
                     (len(restore_ini_txt_date_month) == 2 or len(restore_ini_txt_date_month) == 1) and \
                      2022 < int(restore_ini_txt_date_year) < 2040 and \
                       (1 <= int(restore_ini_txt_date_month) <= 12):
                    restore_input_date_ok = True
                else:
                    restore_input_date_ok = False
            except:
                restore_input_date_ok = False

            print(1.1)
            try:
                if not restore_input_date_ok or restore_ini_txt_date_year_int > cur_date_year or \
                    (restore_ini_txt_date_year_int == cur_date_year and
                     restore_ini_txt_date_month_int > cur_date_month):
                    with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"Restoring. Input data is off.\n")
                else:
                    restore_continue = False
                    # Searching for an object with given name.
                    for object_com_params in objects_com_params:
                        # If there is an object with required name, get it common params.
                        if object_com_params[0][4] == restore_ini_txt_name:
                            restore_object_com_params = object_com_params
                            restore_continue = True
                            break
                    print(1.2)
                    if restore_continue:
                        restore_continue = False
                        # Get number of days in restoring report month.
                        restore_ini_txt_month_days = monthrange(int(restore_ini_txt_date_year),
                                                                int(restore_ini_txt_date_month))[1]
                        # Parametrise MBTCP connection.
                        restore_modbus_client = ModbusClient(host=restore_object_com_params[2][0],
                                                             port=restore_object_com_params[2][1],
                                                             unit_id=restore_object_com_params[2][2],
                                                             timeout=restore_object_com_params[2][3],
                                                             auto_open=True,
                                                             auto_close=True)
                        # Poll parameters for object.
                        restore_list2 = []
                        restore_list2_real = []
                        for param_num in range(restore_object_com_params[1]):
                            # Write year, month, parameter number.
                            restore_w_request = \
                                restore_modbus_client.write_multiple_registers(0,
                                                                               [int(restore_ini_txt_date_year),
                                                                                int(restore_ini_txt_date_month),
                                                                                param_num + 1])
                            # Read parameter values for month.
                            restore_r_request = restore_modbus_client.read_input_registers(0,
                                                                                           restore_ini_txt_month_days * 2)
                            # Draw a list of registers.
                            restore_list2.append(restore_r_request)
                        print(1.3)
                        # If polls succeeded then format data.
                        if restore_modbus_client.last_error == 0:
                            for restore_list in restore_list2:
                                # Conversion list of 2int16 to list of long32.
                                restore_r_request_long_32 = utils.word_list_to_long(restore_list, big_endian=True)
                                # Conversion of list of long32 to list of real.
                                restore_r_request_list_real = []
                                for elem in restore_r_request_long_32:
                                    restore_r_request_list_real.append(utils.decode_ieee(int(elem)))
                                # Add every list of reals to list2 of reals.
                                restore_list2_real.append(restore_r_request_list_real)
                            # Check if the list2 filled fully. It must have 7 (2 + 5) rows in total.
                            restore_list_real_empty = []
                            for i in range(restore_ini_txt_month_days):
                                restore_list_real_empty.append(0.0)
                            # Add empty lists if not filled.
                            for i in range(5 - len(restore_list2_real)):
                                restore_list2_real.append(restore_list_real_empty)
                            print(1.4)
                            # Create list for restoring data.
                            restore_data = []
                            # Add three lines of common object data to list.
                            restore_data.append(restore_object_com_params[3][0])
                            restore_data.append(restore_object_com_params[3][1])
                            restore_data.append(restore_object_com_params[3][2])
                            # Filling of list with polled values.
                            for day in range(restore_ini_txt_month_days):
                                # Add new line in data list.
                                restore_data.append([])
                                # Add the first two meaning in line.
                                restore_data[len(restore_data) - 1].append(day + 1)
                                restore_data[len(restore_data) - 1].append(f"{restore_ini_txt_date_year_int}."
                                                                           f"{restore_ini_txt_date_month_int:02}."
                                                                           f"{(day + 1):02}")
                                # Add meanings of polled values.
                                for restore_list_real in restore_list2_real:
                                    restore_data[len(restore_data) - 1].append(format(restore_list_real[day], '.1f'))
                            print(1.5)
                            # Restore ini.txt file.
                            restore_path = rf"..\reports\{restore_ini_txt_name}\{restore_ini_txt_name}_ini_" \
                                           rf"{restore_ini_txt_date_year_int}_{restore_ini_txt_date_month_int:02}.txt"
                            with open(restore_path, "w", encoding="UTF-8") as restore_ini_txt:
                                for line in restore_data:
                                    for item in line:
                                        restore_ini_txt.write(f"{item};")
                                    restore_ini_txt.write("\n")
                            print(1.6)
                        else:
                            print(f"Restore. MBTCP error while polling")
                            with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                                logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                           f"Restore. MBTCP error while polling.\n")
                    else:
                        print(f"Restore. Object name {restore_ini_txt_name} doesn't exist")
                        with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                            logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                       f"Restore. Object name {restore_ini_txt_name} doesn't exist.\n")
            except:
                print("Restoring. Exception while restoring")
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               f"Restoring. Exception while restoring.\n")


        time.sleep(1.0)