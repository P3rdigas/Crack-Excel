import os
from backend.zipHandler import convert_excel_to_zip, crack_excel, create_unprotected_file

# Class for the extensions supported
class ExcelExtensions:
    CONST_XLSM_EXTENSION = ".xlsm"
    CONST_XLSX_EXTENSION = ".xlsx"

# Get information of the given excel file
source_folder_path = os.getcwd()
excel_file = os.path.basename("Livro2.xlsx") # TODO: Change for soure_folder_path
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
        create_unprotected_file(cracked_zip_file, excel_name, excel_extension)
    case ExcelExtensions.CONST_XLSX_EXTENSION:
        print("You are trying to crack a {} excel file".format(excel_extension))
        zip_file = convert_excel_to_zip(excel_file)
        cracked_zip_file = crack_excel(zip_file)
        create_unprotected_file(cracked_zip_file, excel_name, excel_extension)
    case _:
        print("This file is not supported yet. Files supported .xlxs and .xlsm")
