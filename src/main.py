from pyModbusTCP.client import ModbusClient
from pyModbusTCP import utils
from fpdf import FPDF
import time
import datetime
import os
import subprocess
from typing import Any
from threading import Thread
from calendar import monthrange
import math

# Create class PDF based on original PDF class from FPDF library
class PDF2(FPDF):

    # Method creating report as a pdf document
    def create_report(self, datetime: datetime, margins: tuple, table_data: list[list[Any, ...], ...], num_params: int):
        if type(table_data) != list:
            # data var includes table head and table contents
            data = [['Data', 'is', 'wrong'], ['0', '0', '0']]
        elif len(table_data) <= 1:
            data = [['Data', 'is', 'wrong'], ['0', '0', '0']]
        else:
            data = [[str(el) for el in row] for row in table_data]
        report_head = data.pop(0)
        # Set minimal width for columns in table using ".rjust" string method
        for i in range(len(data)):
            for j in range(len(report_head)):

                if j == 0:
                    data[i][j] = data[i][j].rjust(3, ' ')
                elif j == 1:
                    data[i][j] = data[i][j].rjust(11, ' ')
                elif j > 1:
                    data[i][j] = data[i][j].rjust(13, ' ')
        report_content = data

        # Formatting pdf document
        left_margin, top_margin, right_margin = 5, 5, 5
        row_height = 5
        font_size = 8
        self.set_margins(left_margin, top_margin, right_margin)
        cur_x, cur_y = left_margin, top_margin
        self.set_xy(cur_x, cur_y)

        # Download the font supporting Unicode and set it. *It has to be added before used
        self.add_font("font_1", "", r"..\etc\font\ARIALUNI.ttf")

        # Create the first page
        self.add_page()

        # Filling in general information
        # First line with current date on the left and company name on the right
        self.set_font("font_1", size=14)
        self.cell(w=(self.w - (margins[0] + abs(margins[2]))) / 2, h=self.font_size, txt="{:%Y.%m.%d}".format(datetime),
                  align="L")
        self.cell(w=(self.w - (margins[0] + abs(margins[2]))) / 2, h=self.font_size, txt='ООО "ПромЭкоВод"', align="R",
                  new_x="LMARGIN",
                  new_y="NEXT")
        # Padding
        self.cell(w=self.w, h=10, new_x="LMARGIN", new_y="NEXT")
        # Site name
        self.set_font("font_1", size=20)
        self.cell(self.w - 20, self.font_size * 1.5, txt=table_data[0][1], align="C", new_x="LEFT",
                  new_y="NEXT")
        # Table title
        self.set_font("font_1", size=14)
        self.cell(self.w - 20, self.font_size * 2, txt=f"Отчет по суточным расходам воды и электроэнергии",
                  align="C", new_x="LEFT", new_y="NEXT")
        self.cell(w=self.w, h=2, new_x="LMARGIN", new_y="NEXT")
        cur_x = self.get_x()
        cur_y = self.get_y()
        print(cur_x, cur_y)

        # Draw a table
        # Find widths of columns in table using specified font
        self.set_font("font_1", size=10)

        # Minimum column width is 20, else the meaning of a number of lines is rough
        # col_width equals to a width of the longest meaning in a respective column of the table
        report_elem_width = list()
        for i in range(len(report_content)):
            report_elem_width.append([])
            for j in range(len(report_head)):
                report_elem_width[i].append(self.get_string_width(report_content[i][j]) + 2)
        # Find the longest lines in columns and assign these values to head titles length
        # Reverse massive of contents length (columns into rows)
        report_elem_width_rev = [[0 for j in range(len(report_content))] for i in range(len(report_head))]
        for i in range(len(report_content)):
            for j in range(len(report_head)):
                report_elem_width_rev[j][i] = report_elem_width[i][j]
        report_head_width = list()
        for i in range(len(report_head)):
            # Indents from edges of the cell equal 2
            report_head_width.append(max(report_elem_width_rev[i]) + 2)
        # Find how many lines head titles take
        report_head_lines_num = list()
        report_head_len = list()
        for i in range(len(report_head)):
            report_head_len.append(self.get_string_width(report_head[i]))
            # How many lines a string takes
            if report_head_len[i] > 0 and report_head_len[i] > report_head_width[i] - 2 and report_head_width[i] > 2:
                report_head_lines_num.append(math.ceil(report_head_len[i] / (report_head_width[i] - 2)) + 2)
            else:
                report_head_lines_num.append(3)
        # Draw a table head
        # Calculate how many Line Feed characters are used in every head cell relating to the highest head cell
        for i in range(len(report_head)):
            self.multi_cell(w=report_head_width[i],
                            h=row_height,
                            txt="\n" + report_head[i] + "\n" * (
                                        max(report_head_lines_num) - report_head_lines_num[i] + 2),
                            border=1,
                            align='C')
            cur_x += report_head_width[i]
            self.set_xy(cur_x, cur_y)
        cur_x = left_margin
        # Make top indent according to number of lines in a head cell
        cur_y = self.get_y() + row_height * max(report_head_lines_num)
        self.set_xy(cur_x, cur_y)
        for i in range(len(report_content)):
            for j in range(len(report_head)):
                self.multi_cell(w=report_head_width[j], h=row_height, txt=report_content[i][j], border=1, align='C')
                cur_x += report_head_width[j]
                self.set_xy(cur_x, cur_y)
            cur_x = left_margin
            cur_y += row_height
            self.set_xy(cur_x, cur_y)
        self.output("new_pdf.pdf")

    # Override footer method so that it numerates pages
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


# If ini.txt file doesn't exist, then create it for report month filling lines by "0.0".
def create_ini_txt(ini_txt_path: str, object: list, month_day_num: int, year_num: int, month_num: int) -> str:
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
        for day in range(month_day_num):
            ini_txt_file.write(f"{day + 1};")
            ini_txt_file.write(f"{day + 1:02}.{month_num:02}.{year_num};")
            for item in range(8):
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
pdf_left_margin = 5
pdf_top_margin = 10
pdf_right_margin = -5
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


# VZU Borodinsky (vboro). Isn't used.

vboro_num_params = 5                            # Number of parameters of object.
# Information about a project for ini.txt.
vboro_obj_params = ("VZU Borodinsky",
                    "ВЗУ Бородинский",
                    r"x:\Отчетность по работе станций\ВЗУ Бородинский",
                    r"..\reports\vzu_borodinsky",
                    "vzu_borodinsky",
                    "",
                    "",
                    "",
                    "",
                    "")
# The header of pdf report.
vboro_pdf_table_head = (r"\n№\n\n\n",
                        r"\nДата/время\n\n\n",
                        r"\nВоды за сутки, м3\n\n",
                        r"\nВоды всего, м3\n\n",
                        r"\nЭнергии за сутки, кВт\n\n",
                        r"\nЭнергии всего, кВт\n\n",
                        r"\nЭнерг. на куб, кВт/м3\n\n"
                        r"",
                        r"",
                        r"")
# Set the mask for pdf table columns.
vboro_pdf_table_mask = ("00000",
                        "0000.00.00  00:00",
                        "0000000.0",
                        "0000000.0",
                        "0000000.0",
                        "0000000.0",
                        "0000000.00",
                        "",
                        "",
                        "")
# List of ini.txt parameters
vboro_ini_txt_params = [vboro_obj_params, vboro_pdf_table_head, vboro_pdf_table_mask]
# MBTCP parameters
vboro_slave_address = '192.168.245.50'
vboro_port = 503
vboro_unit_id = 1
vboro_timeout = 7.0
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
                    "",
                    "",
                    "",
                    "")
# The header of pdf report.
vtepl_pdf_table_head = (r"\n№\n\n\n",
                        r"\nДата/время\n\n\n",
                        r"\nРасход на входе 1, м3\n\n",
                        r"\nРасход на входе 2, м3\n\n",
                        r"\nРасход на выходе, м3\n\n",
                        "",
                        "",
                        "",
                        "",
                        "")
# Set the mask for pdf table columns.
vtepl_pdf_table_mask = ("00000",
                        "0000.00.00  00:00",
                        "0000000000000.0",
                        "0000000000000.0",
                        "0000000000000.0",
                        "",
                        "",
                        "",
                        "",
                        "")
# Parameters of ini.txt.
vtepl_ini_txt_params = [vtepl_obj_params, vtepl_pdf_table_head, vtepl_pdf_table_mask]
# MBTCP parameters.
vtepl_slave_address = '192.168.241.50'
vtepl_port = 503
vtepl_unit_id = 1
vtepl_timeout = 7.0
vtepl_mbtcp_params = (vtepl_slave_address, vtepl_port, vtepl_unit_id, vtepl_timeout)
# List of vtepl object parameters.
vtepl_common_params = [vtepl_obj_params, vtepl_num_params, vtepl_mbtcp_params, vtepl_ini_txt_params]    # Common parameters for object.


# VZU Chern (vchrn).

vchrn_num_params = 8                            # Number of parameters of object.
# Information about project for ini.txt.
vchrn_obj_params = ("VZU Chern",
                    "ВЗУ Чернь",
                    r"c:\Отчетность по работе станций\ВЗУ Чернь",
                    r"..\reports\vzu_chern",
                    "vzu_chern",
                    "",
                    "",
                    "",
                    "",
                    "")
# The header of pdf report.
vchrn_pdf_table_head = (r"\n№\n\n\n\n",
                        r"\nДата/время\n\n\n\n",
                        r"\nРасход на входе 1, м3\n\n",
                        r"\nРасход на входе 2, м3\n\n",
                        r"\nРасход на входе, м3\n\n\n",
                        r"\nРасход на промыв., м3\n\n",
                        r"\nРасход на выходе 1, м3\n\n",
                        r"\nРасход на выходе 2, м3\n\n",
                        r"\nРасход на выходе, м3\n\n\n",
                        r"\nРасход энергии, КВтч\n\n")
# Set the mask for pdf table columns.
vchrn_pdf_table_mask = ("000",
                        "0000.00.00",
                        "00000000.0",
                        "00000000.0",
                        "00000000.0",
                        "00000000.0",
                        "000000000.0",
                        "000000000.0",
                        "000000000.0",
                        "000000000.0")
# Parameters of ini.txt.
vchrn_ini_txt_params = [vchrn_obj_params, vchrn_pdf_table_head, vchrn_pdf_table_mask]
# MBTCP parameters.
vchrn_slave_address = '192.168.236.50'
vchrn_port = 503
vchrn_unit_id = 1
vchrn_timeout = 7.0
vchrn_mbtcp_params = (vchrn_slave_address, vchrn_port, vchrn_unit_id, vchrn_timeout)
# List of vtepl object parameters.
vchrn_common_params = [vchrn_obj_params, vchrn_num_params, vchrn_mbtcp_params, vchrn_ini_txt_params]    # Common parameters for object.


# KOS Makarovo  (kmkrv).

kmkrv_num_params = 3                            # Number of parameters of object.
# Information about project for ini.txt.
kmkrv_obj_params = ("KOS Makarovo",
                    "КОС Макарово",
                    r"c:\Отчетность по работе станций\КОС Макарово",
                    r"..\reports\kos_makarovo",
                    "kos_makarovo",
                    "",
                    "",
                    "",
                    "",
                    "")
# The header of pdf report.
kmkrv_pdf_table_head = (r"\n№\n\n\n",
                        r"\nДата/время\n\n\n",
                        r"\nРасход на входе 1, м3\n\n",
                        r"\nРасход на входе 2, м3\n\n",
                        r"\nРасход на выходе, м3\n\n",
                        "",
                        "",
                        "",
                        "",
                        "")
# Set the mask for pdf table columns.
kmkrv_pdf_table_mask = ("00000",
                        "0000.00.00  00:00",
                        "0000000000000.0",
                        "0000000000000.0",
                        "0000000000000.0",
                        "",
                        "",
                        "",
                        "",
                        "")
# Parameters of ini.txt.
kmkrv_ini_txt_params = [kmkrv_obj_params, kmkrv_pdf_table_head, kmkrv_pdf_table_mask]
# MBTCP parameters.
kmkrv_slave_address = '192.168.239.50'
kmkrv_port = 503
kmkrv_unit_id = 1
kmkrv_timeout = 7.0
kmkrv_mbtcp_params = (kmkrv_slave_address, kmkrv_port, kmkrv_unit_id, kmkrv_timeout)
# List of kmkrv object parameters.
kmkrv_common_params = [kmkrv_obj_params, kmkrv_num_params, kmkrv_mbtcp_params, kmkrv_ini_txt_params]    # Common parameters for object.


# VZU Nord  (vnord).

vnord_num_params = 5                            # Number of parameters of object.
# Information about project for ini.txt.
vnord_obj_params = ("VZU Nord",
                    "ВЗУ Норд",
                    r"c:\Отчетность по работе станций\ВЗУ Норд",
                    r"..\reports\vzu_nord",
                    "vzu_nord",
                    "",
                    "",
                    "",
                    "",
                    "")
# The header of pdf report.
vnord_pdf_table_head = (r"\n№\n\n\n\n",
                        r"\nДата/время\n\n\n\n",
                        r"\nОсмос, м3\n\n\n\n",
                        r"\nРЧВ, м3\n\n\n\n",
                        r"\nВход, м3\n\n\n\n",
                        r"\nВыход 1, м3\n\n\n\n",
                        r"\nВыход 2, м3\n\n\n\n",
                        r"",
                        r"",
                        r"")
# Set the mask for pdf table columns.
vnord_pdf_table_mask = ("00000",
                        "0000.00.00  00:00",
                        "00000000000.0",
                        "00000000000.0",
                        "00000000000.0",
                        "00000000000.0",
                        "00000000000.0",
                        "",
                        "",
                        "")
# Parameters of ini.txt.
vnord_ini_txt_params = [vnord_obj_params, vnord_pdf_table_head, vnord_pdf_table_mask]
# MBTCP parameters.
vnord_slave_address = '192.168.238.50'
vnord_port = 503
vnord_unit_id = 1
vnord_timeout = 7.0
vnord_mbtcp_params = (vnord_slave_address, vnord_port, vnord_unit_id, vnord_timeout)
# List of kmkrv object parameters.
vnord_common_params = [vnord_obj_params, vnord_num_params, vnord_mbtcp_params, vnord_ini_txt_params]    # Common parameters for object.


# List of objects.
objects_com_params = [vchrn_common_params, vtepl_common_params, vnord_common_params]

# Check existence of ini.txt file.
# for object in objects_com_params:
#     Get path to object ini.txt.
    # object_ini_txt_path = rf"{object[0][3]}/{object[0][4]}_ini_{cur_date_year}_{cur_date_month:02}.txt"
    # If exists, then check number of lines.
    # if os.path.isfile(object_ini_txt_path):
    #     with open(f"{object_ini_txt_path}", "r", encoding="UTF-8") as object_ini_txt:
    #         ini_txt_file_list_lines = object_ini_txt.readlines()
    #         If file has 3 or more lines it's ok.
            # if len(ini_txt_file_list_lines) >= 3:
            #     with open("..\logs.txt", "a", encoding="UTF-8") as logs:
            #         logs.write(
            #             f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}{object[0]}_ini.txt exists\n")
            # else:
            #     with open("..\logs.txt", "a", encoding="UTF-8") as logs:
            #         logs.write(
            #             f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}{object[0]}_ini.txt is off\n")
    # Else, renew file.
    # else:
    #     create_ini_txt(object=object, ini_txt_path=object_ini_txt_path, day_prev_num=cur_date_day - 1)

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
        # if cur_time_hour == 0 and cur_time_min == 0 and cur_time_sec == 5:
        if cur_time_min == 53 and cur_time_sec == 10:
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
                # Get number of days in restoring report month.
                object_ini_txt_month_days = monthrange(int(report_last_year), int(report_last_month))[1]
                # Set path to object's ini.txt file.
                object_ini_txt_path = f"{object[0][3]}/{object[0][4]}_ini_{report_last_year}_{report_last_month:02}.txt"
                # Check existence of ini.txt file. If exists then move on, else create new ini.txt file.
                if not os.path.isfile(object_ini_txt_path):
                    create_ini_txt(object=object,
                                   ini_txt_path=object_ini_txt_path,
                                   month_day_num=object_ini_txt_month_days,
                                   year_num=report_last_year,
                                   month_num=report_last_month)
                # Parametrise MBTCP connection.
                object_mbtcp_client = ModbusClient(host=object[2][0],
                                                   port=int(object[2][1]),
                                                   unit_id=int(object[2][2]),
                                                   timeout=float(object[2][3]),
                                                   auto_open=True,
                                                   auto_close=True)

                # Poll parameters for object.
                object_list2 = []
                object_list2_real = []
                for param_num in range(object[1]):
                    # Write year, month, parameter number.
                    object_poll_w_status = \
                        object_mbtcp_client.write_multiple_registers(0,
                                                                       [int(report_last_year),
                                                                        int(report_last_month),
                                                                        param_num + 1])
                    # Read parameter values for month.
                    object_poll_r = object_mbtcp_client.read_input_registers(0, object_ini_txt_month_days * 2)
                    # Draw a list of registers.
                    object_list2.append(object_poll_r)

                # If polls are successful then format data.
                if object_mbtcp_client.last_error == 0:
                    for object_list in object_list2:
                        # Conversion list of 2 int16 to list of long32.
                        object_r_request_long_32 = utils.word_list_to_long(object_list, big_endian=True)
                        # Conversion of list of long32 to list of real.
                        object_r_request_list_real = []
                        for elem in object_r_request_long_32:
                            object_r_request_list_real.append(utils.decode_ieee(int(elem)))
                        # Add every list of reals to list2 of reals.
                        object_list2_real.append(object_r_request_list_real)
                    # Check if the list2 filled fully. It must have 10 (2 + 8) rows in total.
                    object_list_real_empty = []
                    for i in range(object_ini_txt_month_days):
                        object_list_real_empty.append(0.0)
                    # Add empty lists if not filled.
                    for i in range(8 - len(object_list2_real)):
                        object_list2_real.append(object_list_real_empty)
                    print(1.4)
                    # Create list2 for restoring data.
                    # Add three lines of common object data to a list.
                    object_data = [object[3][0], object[3][1], object[3][2]]
                    # Filling of a list with polled values.
                    for day in range(object_ini_txt_month_days):
                        # Add new line in a data list.
                        object_data.append([])
                        # Add the first two meaning in a line.
                        object_data[len(object_data) - 1].append(day + 1)
                        object_data[len(object_data) - 1].append(f"{(day + 1):02}.{report_last_month:02}."
                                                                 f"{report_last_year}")
                        # Add meanings of polled values.
                        for object_list_real in object_list2_real:
                            object_data[len(object_data) - 1].append(format(object_list_real[day], '.1f'))
                    print(1.5)
                    # Add sum values at the end of data list (the last line in object_data)
                    object_sum_line = ["","За месяц:"]
                    for i in range(8):
                        object_sum_value_in_month = 0.0
                        for j in range(object_ini_txt_month_days):
                            object_sum_value_in_month = object_sum_value_in_month + object_list2_real[i][j]
                        object_sum_line.append(format(object_sum_value_in_month, '.1f'))
                    object_data.append(object_sum_line)
                    # Write ini.txt file.
                    with open(object_ini_txt_path, "w", encoding="UTF-8") as object_ini_txt:
                        for line in object_data:
                            for item in line:
                                object_ini_txt.write(f"{item};")
                            object_ini_txt.write("\n")
                    print(1.6)
                else:
                    print(f"{object[0][4]}. MBTCP error while polling")
                    with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"{object[0][4]}. MBTCP error while polling.\n")

            # PDF document printing.

            for object in objects_com_params:
                report_pdf = PDF2()
                # Set path to object's ini.txt file.
                object_ini_txt_path = f"{object[0][3]}/{object[0][4]}_ini_{report_last_year}_{report_last_month:02}.txt"
                # print(object_ini_txt_path)
                # Read data from ini.txt and write it into list2.
                object_data_list2 = []
                read_ini_txt(object_ini_txt_path, object_data_list2)
                # print(object_data_list2)
                # Fill in pdf.
                print()
                report_pdf.create_report(datetime=report_last_datetime, margins=pdf_margins,
                                         table_data=object_data_list2, num_params=object[1])
                # Check existence of pdf report directory.
                check_pdf_dir_return = check_pdf_dir(object[3][0])
                pdf_dir_exist = check_pdf_dir_return[0]
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               f"{check_pdf_dir_return[1]}\n")
                # if pdf_dir_exist:
                #     print(rf"x:\Отчетность по работе станций\{object[0][1]}\Отчет по {object[0][1]} от {report_last_year}_{report_last_month:02}.pdf")
                #     report_pdf.output(
                #         rf"x:\Отчетность по работе станций\{object[0][1]}\Отчет по {object[0][1]} от {report_last_year}_{report_last_month:02}.pdf")
                # else:
                #     report_pdf.output(
                #         rf"c:\Отчетность по работе станций\{object[0][1]}\Отчет по {object[0][1]} от {report_last_year}_{report_last_month:02}.pdf")
                report_pdf.output(
                                  rf"c:\Отчетность по работе станций\{object[0][1]}\Отчет по {object[0][1]} от "
                                  rf"{report_last_year}_{report_last_month:02}.pdf")


        # Print a report for specified time if it was deleted.
        if print_start:
            print_start = False
            # Checking input restoring date for adequacy.
            print(2.1)
            try:
                print_ini_txt_date_list = print_ini_txt_date.split("_")
                print_ini_txt_date_year = print_ini_txt_date_list[0]
                print_ini_txt_date_year_int = int(print_ini_txt_date_year)
                print_ini_txt_date_month = print_ini_txt_date_list[1]
                print_ini_txt_date_month_int = int(print_ini_txt_date_month)
                print(2.2)
                if len(print_ini_txt_date_list) == 2 and \
                        len(print_ini_txt_date_year) == 4 and \
                        (len(print_ini_txt_date_month) == 2 or len(print_ini_txt_date_month) == 1) and \
                        2022 < int(print_ini_txt_date_year) < 2040 and \
                        (1 <= int(print_ini_txt_date_month) <= 12):
                    # Path to ini.txt file.
                    print_path = rf"..\reports\{print_ini_txt_name}\{print_ini_txt_name}_ini_{print_ini_txt_date_year_int}_" \
                                 rf"{print_ini_txt_date_month_int:02}.txt"
                    print(2.3)
                    if os.path.isfile(print_path):
                        print(2.4)
                        # Get parameters number of object.
                        print_num_params = 0
                        for object in objects_com_params:
                            if object[3][0][4] == print_ini_txt_name:
                                print(2.5)
                                print(print_ini_txt_name)
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
                        pdf_print = PDF2()
                        # Create report.
                        pdf_print.create_report(datetime=cur_datetime, margins=pdf_margins, table_data=print_data,
                                                num_params=print_num_params)
                        print(2.6)
                        # Check existence of pdf report directory.
                        check_print_pdf_dir_return = check_pdf_dir(print_object_params)
                        print_pdf_dir_exist = check_print_pdf_dir_return[0]
                        with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                            logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                       f"{check_print_pdf_dir_return[1]}\n")
                        # if print_pdf_dir_exist:
                        #     pdf_print.output(
                        #         rf"x:\Отчетность по работе станций\{print_object_params[1]}\Отчет по {print_object_params[1]} "
                        #         rf"от {print_ini_txt_date_year_int}_{print_ini_txt_date_month_int:02}_восстановленный.pdf")
                        # else:
                        #     pdf_print.output(
                        #         rf"c:\Отчетность по работе станций\{print_object_params[1]}\Отчет по {print_object_params[1]} "
                        #         rf"от {print_ini_txt_date_year_int}_{print_ini_txt_date_month_int:02}_восстановленный.pdf")
                        pdf_print.output(
                                rf"c:\Отчетность по работе станций\ВЗУ Чернь\Отчет по ВЗУ Чернь "
                                rf"от {print_ini_txt_date_year_int}_{print_ini_txt_date_month_int:02}_восстановленный.pdf")
                    else:
                        with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                            logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                       f"Printing. File in {print_path} doesn't exist.\n")
            except:
                print(2.7)
                print_check_ok = False
                with open("..\logs.txt", "a", encoding="UTF-8") as logs:
                    logs.write(rf"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               rf"Printing. Exception while checking input data." + "\n")


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
                        # If polls are succeeded then format data.
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
                            # Check if the list2 filled fully.
                            restore_list_real_empty = []
                            # Create empty list.
                            for i in range(restore_ini_txt_month_days):
                                restore_list_real_empty.append(0.0)
                            # Add empty lists at the end if data is not filled.
                            for i in range(8 - len(restore_list2_real)):
                                restore_list2_real.append(restore_list_real_empty)
                            print(1.4)
                            # Create list for restoring data.
                            # Add three lines of common object data to list.
                            restore_data = [restore_object_com_params[3][0], restore_object_com_params[3][1],
                                            restore_object_com_params[3][2]]
                            # Filling of list with polled values.
                            for day in range(restore_ini_txt_month_days):
                                # Add new line in data list.
                                restore_data.append([])
                                # Add the first two meaning in line.
                                restore_data[len(restore_data) - 1].append(day + 1)
                                restore_data[len(restore_data) - 1].append(f"{(day + 1):02}."
                                                                           f"{restore_ini_txt_date_month_int:02}."
                                                                           f"{restore_ini_txt_date_year_int}")
                                # Add meanings of polled values.
                                for restore_list_real in restore_list2_real:
                                    restore_data[len(restore_data) - 1].append(format(restore_list_real[day], '.1f'))
                            print(1.5)
                            # Add sum values at the end of data list (the last line in object_data)
                            restore_sum_line = ["", "За месяц:"]
                            for i in range(8):
                                restore_sum_value_in_month = 0.0
                                for j in range(restore_ini_txt_month_days):
                                    restore_sum_value_in_month = restore_sum_value_in_month + restore_list2_real[i][j]
                                restore_sum_line.append(format(restore_sum_value_in_month, '.1f'))
                            restore_data.append(restore_sum_line)
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
                               f"Restore. Exception while restoring.\n")


        time.sleep(1.1)