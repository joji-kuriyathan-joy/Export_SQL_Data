# Export_SQL_Data
##############################################################
'''
Export_SQL_Data_V1.0                           Feb-02-2024
This python file executes SQL statements and Export Data into EXCEL/CSV
Options to Export with header and with out headers
Execution process is written down to Log file in Log Folder
Exported data in the form of EXCEL/CSV is present in the Output Folder after Execution
If the execution type is EXCEL , then the  single Excel contains sheets of executed query from the query_list
Note:
    Make sure that the sheet_list and csv_filename_list is equal to the query_list in the config.
    Specify the Type of Export in the config
    Specify the Output and Log Folder path in the config
    Give SQL connection details in the config. The DB used is SQL Server
    Change to True / False for with column headers and with out
    Change the delimiter to ',' or '|' while exporting to CSV 
'''
###############################################################
