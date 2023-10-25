import os
from backend.zipHandler import convert_excel_to_zip, crack_excel

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
        zip_file = convert_excel_to_zip(excel_file)
        cracked_zip_file = crack_excel(zip_file)
    case _:
        print("This file is not supported yet. Files supported .xlxs and .xlsm")
