    This program has been created to solve the issue of collecting and printing data gained for a long period of time by ModbusTCP protocol.
The program is a client (futher the Client) which can connect to various devices and poll them one by one receiving several technological parameters (e.g
water volume per day). This requires a program created in your PLC (programmable logic controller) beforehand. You should create in your
PLC some code that provides modbus registers and renewable data rewritten into this registers every midnight, so that the Client polls
these registers after midnight daily and gets required data.
    The way the Client store data from PLCs is saving data as text files into 'reports' folder in the project directory using different folders in the
'reports'. TXT files have the specified format using semicolon delimiter (;) to devide rows to columns. There is a feature implemented in the Client
allowing to create a PDF document from a TXT file.
