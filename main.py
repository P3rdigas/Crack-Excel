import os
import sys
import webbrowser
from PIL import Image
from CTkMenuBar import *
import customtkinter as ctk
from utils import convert_excel_to_zip, crack_excel, create_unprotected_file

# Class for the extensions supported
class ExcelExtensions:
    CONST_XLSM_EXTENSION = ".xlsm"
    CONST_XLSX_EXTENSION = ".xlsx"

# Python colors: https://matplotlib.org/stable/gallery/color/named_colors.html
CONST_MENUBAR_BACKGROUND_COLOR_LIGHT = "grey75"
CONST_MENUBAR_BACKGROUND_COLOR_DARK = "black"
CONST_MENUBAR_TEXT_COLOR_LIGHT = "black"
CONST_MENUBAR_TEXT_COLOR_DARK = "white"
CONST_MENUBAR_HOVER_COLOR_LIGHT = "black"
CONST_MENUBAR_HOVER_COLOR_DARK = "grey25"

CONST_GITHUB_LOGO_LIGHT_PATH = r"assets/icons/github-mark.png"
CONST_GITHUB_LOGO_DARK_PATH = r"assets/icons/github-mark-white.png"

CONST_SOURCE_CODE_URL = "https://github.com/P3rdigas/Crack-Excel"

root = ctk.CTk()

root.title("Crack Excel")
root.iconbitmap('assets/logos/cracked_excel_logo_128x128.ico')
root.geometry("800x600")
root.resizable(width=False, height=False)

# https://github.com/Akascape/CTkMenuBar
menu_bar_bg = CONST_MENUBAR_BACKGROUND_COLOR_DARK
menu_bar_text_color = CONST_MENUBAR_TEXT_COLOR_DARK
menu_bar_hover_color = CONST_MENUBAR_HOVER_COLOR_DARK
toolbar = CTkMenuBar(master=root, bg_color=menu_bar_bg, border_width=1)
file_button = toolbar.add_cascade("File", text_color=menu_bar_text_color, hover_color=menu_bar_hover_color)
settings_button = toolbar.add_cascade("Settings", text_color=menu_bar_text_color, hover_color=menu_bar_hover_color)
about_button = toolbar.add_cascade("About", text_color=menu_bar_text_color, hover_color=menu_bar_hover_color)

# File Button Functions
file_button_dropdown = CustomDropdownMenu(widget=file_button, corner_radius=5)
file_button_dropdown.add_option(option="Import")
file_button_dropdown.add_option(option="Export")
file_button_dropdown.add_option(option="Exit", command=root.destroy)

# About Button Functions
def open_browser():
    webbrowser.open_new(CONST_SOURCE_CODE_URL)

image = ctk.CTkImage(Image.open(CONST_GITHUB_LOGO_DARK_PATH)) 
about_button_dropdown = CustomDropdownMenu(widget=about_button, corner_radius=5)
about_button_dropdown.add_option(option="Source Code", image=image, command=open_browser)

# TODO: Button to change appearance mode (começar em system, deixar mudar para dark ou light)
# window._set_appearance_mode("dark")

root.mainloop()



# if len(sys.argv) != 2:
#     print("Usage: py main.py filename.extension")
#     sys.exit(1)

# excel_file = sys.argv[1]

# # Get information of the given excel file
# source_folder_path = os.getcwd()
# excel_name, excel_extension = os.path.splitext(excel_file)

# #
# # Excel file extensions
# # Link: https://support.microsoft.com/en-au/office/file-formats-that-are-supported-in-excel-0943ff2c-6014-4e8d-aaea-b83d51d46247
# #
# # Treats each extension differently if needed
# match excel_extension:
#     case ExcelExtensions.CONST_XLSM_EXTENSION:
#         print("You are trying to crack a {} excel file".format(excel_extension))
#         zip_file = convert_excel_to_zip(excel_file)
#         cracked_zip_file = crack_excel(zip_file)
#         create_unprotected_file(cracked_zip_file, excel_name, excel_extension)
#     case ExcelExtensions.CONST_XLSX_EXTENSION:
#         print("You are trying to crack a {} excel file".format(excel_extension))
#         zip_file = convert_excel_to_zip(excel_file)
#         cracked_zip_file = crack_excel(zip_file)
#         create_unprotected_file(cracked_zip_file, excel_name, excel_extension)
#     case _:
#         print("This file is not supported yet. Files supported .xlxs and .xlsm")
