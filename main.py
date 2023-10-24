import os

from backend.zipHandler import *

# Class for the extensions supported
class ExcelExtensions:
    CONST_XLSM_EXTENSION = ".xlsm"

# Get information of the given excel file
source_folder_path = os.getcwd()
excel_file = os.path.basename("input.xlsm") # TODO: Change for soure_folder_path
excel_name, excel_extension = os.path.splitext(excel_file)

#
# Excel file extensions
# Link: https://support.microsoft.com/en-au/office/file-formats-that-are-supported-in-excel-0943ff2c-6014-4e8d-aaea-b83d51d46247
#
# Treats each extension differently if needed
match excel_extension:
    case ExcelExtensions.CONST_XLSM_EXTENSION:
        print("You are trying to crack a {} excel file".format(excel_extension))
        convert_excel_to_zip(source_folder_path, excel_file, excel_name)
    case _:
        print("This file is not supported yet. Files supported .xlxs and .xlsm")



# #
# # Class for the extensions supported
# class ExcelExtensions:
#     CONST_XLXS_EXTENSION = ".xlxs"
#     CONST_XLSM_EXTENSION = ".xlsm"

# #
# # Function that handle excel files with .xlxs extension
# def handle_xlxs_extention(excel_name, excel_extension, excel_file_path, source_folder_path):
#     print("You are trying to crack a {} excel file".format(excel_extension))

# #
# # Function that handle excel files with .xlsm extension
# def handle_xlsm_extention(excel_name, excel_extension, excel_file_path, source_folder_path):
#     print("You are trying to crack a {} excel file".format(excel_extension))
#     convert_to_zip(excel_name, excel_file_path, source_folder_path)


# #
# # Gets the path of the working directory of the main.py file
# source_folder_path = os.getcwd()

# zip_file_name = "input.zip"
# zip_file_path = os.path.join(source_folder_path, zip_file_name)

# if(os.path.exists(zip_file_path)):
#     open_zip_file()
# else:
#     #
#     # Gets the name and the extension of the excel file
#     excel_file_name = "input.xlsm"
#     excel_name, excel_extension = os.path.splitext(excel_file_name)
#     excel_file_path = os.path.join(source_folder_path, excel_file_name)
    
#     #
#     # Excel file extensions
#     # Link: https://support.microsoft.com/en-au/office/file-formats-that-are-supported-in-excel-0943ff2c-6014-4e8d-aaea-b83d51d46247
#     #
#     # Treats each extension differently if needed
#     match excel_extension:
#         case ExcelExtensions.CONST_XLXS_EXTENSION:
#             handle_xlxs_extention(excel_name, excel_extension, excel_file_path, source_folder_path)
#         case ExcelExtensions.CONST_XLSM_EXTENSION:
#             handle_xlsm_extention(excel_name, excel_extension, excel_file_path, source_folder_path)
#         case _:
#             print("This file is not supported yet. Files supported .xlxs and .xlsm")

