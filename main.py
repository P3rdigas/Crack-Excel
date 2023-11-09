import os
import sys
import customtkinter as ctk
from utils import convert_excel_to_zip, crack_excel, create_unprotected_file

# Class for the extensions supported
class ExcelExtensions:
    CONST_XLSM_EXTENSION = ".xlsm"
    CONST_XLSX_EXTENSION = ".xlsx"

window = ctk.CTk()

window.title("Crack Excel")
window.iconbitmap('assets/cracked_excel_logo_128x128.ico')
window.geometry("800x600")
window.resizable(width=False, height=False)

# TODO: Button to change appearance mode (começar em system, deixar mudar para dark ou light)
# window._set_appearance_mode("dark")


window.mainloop()








if len(sys.argv) != 2:
    print("Usage: py main.py filename.extension")
    sys.exit(1)

excel_file = sys.argv[1]

# Get information of the given excel file
source_folder_path = os.getcwd()
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
