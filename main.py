import os
import sys
import webbrowser
import configparser
from PIL import Image
from CTkMenuBar import *
import customtkinter as ctk
from utils import convert_excel_to_zip, crack_excel, create_unprotected_file

# Class for the extensions supported
class ExcelExtensions:
    CONST_XLSM_EXTENSION = ".xlsm"
    CONST_XLSX_EXTENSION = ".xlsx"

# Python colors: https://matplotlib.org/stable/gallery/color/named_colors.html or https://www.discogcodingacademy.com/turtle-colours
MENUBAR_BACKGROUND_COLOR_LIGHT = "white"
MENUBAR_BACKGROUND_COLOR_DARK = "black"

DROPDOWN_BACKGROUND_COLOR_LIGHT = "white"
DROPDOWN_BACKGROUND_COLOR_DARK = "grey20"

TEXT_COLOR_LIGHT = "black"
TEXT_COLOR_DARK = "white"
HOVER_COLOR_LIGHT = "light gray"
HOVER_COLOR_DARK = "grey25"

GITHUB_LOGO_LIGHT_PATH = r"assets/icons/github-mark.png"
GITHUB_LOGO_DARK_PATH = r"assets/icons/github-mark-white.png"

SOURCE_CODE_URL = "https://github.com/P3rdigas/Crack-Excel"

root = ctk.CTk()

root.title("Crack Excel")
root.iconbitmap('assets/logos/cracked_excel_logo_128x128.ico')
root.geometry("800x600")
root.resizable(width=False, height=False)

# https://github.com/Akascape/CTkMenuBar
menu_bar_bg = None
menu_bar_text_color = None
menu_bar_hover_color = None
dropdown_bg_color = None
dropdown_text_color = None
dropdown_hover_color = None
github_image = None

def load_init_light_mode():
    global menu_bar_bg, menu_bar_text_color, menu_bar_hover_color, dropdown_bg_color, dropdown_text_color, dropdown_hover_color, github_image

    menu_bar_bg = MENUBAR_BACKGROUND_COLOR_LIGHT
    menu_bar_text_color = TEXT_COLOR_LIGHT
    menu_bar_hover_color = HOVER_COLOR_LIGHT
    dropdown_bg_color = DROPDOWN_BACKGROUND_COLOR_LIGHT
    dropdown_text_color = TEXT_COLOR_LIGHT
    dropdown_hover_color = HOVER_COLOR_LIGHT

    github_image = ctk.CTkImage(Image.open(GITHUB_LOGO_LIGHT_PATH))

def load_init_dark_mode():
    global menu_bar_bg, menu_bar_text_color, menu_bar_hover_color, dropdown_bg_color, dropdown_text_color, dropdown_hover_color, github_image

    menu_bar_bg = MENUBAR_BACKGROUND_COLOR_DARK
    menu_bar_text_color = TEXT_COLOR_DARK
    menu_bar_hover_color = HOVER_COLOR_DARK
    dropdown_bg_color = DROPDOWN_BACKGROUND_COLOR_DARK
    dropdown_text_color = TEXT_COLOR_DARK
    dropdown_hover_color = HOVER_COLOR_DARK

    github_image = ctk.CTkImage(Image.open(GITHUB_LOGO_DARK_PATH))

def load_init_system_mode():
    root._set_appearance_mode("system")
    appearance_mode = root._get_appearance_mode()

    if appearance_mode == "light":
        load_init_light_mode()
    else:
        load_init_dark_mode()

# Loading config file
config_file = 'config.ini'

config = configparser.ConfigParser()

if os.path.isfile(config_file):
    config.read(config_file)
    theme = config.get('Settings', 'theme')

    if theme == "system":
        load_init_system_mode()
    elif theme == "light":
        root._set_appearance_mode("light")
        load_init_light_mode()
    else:
        root._set_appearance_mode("dark")
        load_init_dark_mode()
else:
    config['Settings'] = {'theme': 'system'}
    with open(config_file, 'w') as configfile:
        config.write(configfile)
    
    load_init_system_mode()

toolbar = CTkMenuBar(master=root, bg_color=menu_bar_bg)
file_button = toolbar.add_cascade("File", text_color=menu_bar_text_color, hover_color=menu_bar_hover_color)
settings_button = toolbar.add_cascade("Settings", text_color=menu_bar_text_color, hover_color=menu_bar_hover_color)
about_button = toolbar.add_cascade("About", text_color=menu_bar_text_color, hover_color=menu_bar_hover_color)

# File Button Functions
file_button_dropdown = CustomDropdownMenu(widget=file_button, corner_radius=0, bg_color=dropdown_bg_color, text_color=dropdown_text_color, hover_color=dropdown_hover_color)
file_button_dropdown.add_option(option="Import")
file_button_dropdown.add_option(option="Export")
file_button_dropdown.add_option(option="Exit", command=root.destroy)

# Settings Button Functions
def load_images(text, mode):
    match text:
        case "Source Code":
            if mode == "light":
                return ctk.CTkImage(Image.open(GITHUB_LOGO_LIGHT_PATH))
            else:
                return ctk.CTkImage(Image.open(GITHUB_LOGO_DARK_PATH))

def updated_dropdown(mode, widget, original_dropdown, bg_color, text_color, hover_color):
    new_dropdown = CustomDropdownMenu(
        widget=widget,
        corner_radius=0,
        bg_color=bg_color,
        text_color=text_color,
        hover_color=hover_color
    )

    for option in original_dropdown._options_list:
        if isinstance(option, dropdown_menu._CDMSubmenuButton):
            submenu_copy = new_dropdown.add_submenu(option.cget("text"), bg_color=bg_color)
            for sub_option in option.submenu._options_list:
                image = None

                if hasattr(sub_option, '_image') and sub_option._image is not None:
                    image = load_images(option.cget("text"), mode)
                    
                submenu_copy.add_option(
                    option=sub_option.cget("text"),
                    image = image,
                    command=sub_option._command
                )
        else:
            image = None
            
            if hasattr(option, '_image') and option._image is not None:
                image = load_images(option.cget("text"), mode)

            new_dropdown.add_option(
                option=option.cget("text"),
                image = image,
                command=option._command
            )

    return new_dropdown

def configure_light_mode():
    global file_button_dropdown, settings_button_dropdown, about_button_dropdown

    root._set_appearance_mode("light")

    toolbar.configure(bg_color=MENUBAR_BACKGROUND_COLOR_LIGHT)
    file_button.configure(text_color=TEXT_COLOR_LIGHT, hover_color=HOVER_COLOR_LIGHT)
    settings_button.configure(text_color=TEXT_COLOR_LIGHT, hover_color=HOVER_COLOR_LIGHT)
    about_button.configure(text_color=TEXT_COLOR_LIGHT, hover_color=HOVER_COLOR_LIGHT)

    file_button_dropdown = updated_dropdown("light", file_button, file_button_dropdown, DROPDOWN_BACKGROUND_COLOR_LIGHT, TEXT_COLOR_LIGHT, HOVER_COLOR_LIGHT)
    settings_button_dropdown = updated_dropdown("light", settings_button, settings_button_dropdown, DROPDOWN_BACKGROUND_COLOR_LIGHT, TEXT_COLOR_LIGHT, HOVER_COLOR_LIGHT)
    about_button_dropdown = updated_dropdown("light", about_button, about_button_dropdown, DROPDOWN_BACKGROUND_COLOR_LIGHT, TEXT_COLOR_LIGHT, HOVER_COLOR_LIGHT)

def load_light_mode():
    configure_light_mode()
    
    config.set('Settings', 'theme', 'light')

    with open(config_file, 'w') as configfile:
        config.write(configfile)

def configure_dark_mode():
    global file_button_dropdown, settings_button_dropdown, about_button_dropdown
    
    root._set_appearance_mode("dark")

    toolbar.configure(bg_color=MENUBAR_BACKGROUND_COLOR_DARK)
    file_button.configure(text_color=TEXT_COLOR_DARK, hover_color=HOVER_COLOR_DARK)
    settings_button.configure(text_color=TEXT_COLOR_DARK, hover_color=HOVER_COLOR_DARK)
    about_button.configure(text_color=TEXT_COLOR_DARK, hover_color=HOVER_COLOR_DARK)

    file_button_dropdown = updated_dropdown("dark", file_button, file_button_dropdown, DROPDOWN_BACKGROUND_COLOR_DARK, TEXT_COLOR_DARK, HOVER_COLOR_DARK)
    settings_button_dropdown = updated_dropdown("dark", settings_button, settings_button_dropdown, DROPDOWN_BACKGROUND_COLOR_DARK, TEXT_COLOR_DARK, HOVER_COLOR_DARK)
    about_button_dropdown = updated_dropdown("dark", about_button, about_button_dropdown, DROPDOWN_BACKGROUND_COLOR_DARK, TEXT_COLOR_DARK, HOVER_COLOR_DARK)

def load_dark_mode():
    configure_dark_mode()

    config.set('Settings', 'theme', 'dark')

    with open(config_file, 'w') as configfile:
        config.write(configfile)

def load_system_mode():
    # TODO: Crash the app in the next line
    # root._set_appearance_mode("system")
    # appearance_mode = root._get_appearance_mode()

    # if appearance_mode == "light":
    #     configure_light_mode()
    # else:
    #     configure_dark_mode()

    config.set('Settings', 'theme', 'system')

    with open(config_file, 'w') as configfile:
        config.write(configfile)

settings_button_dropdown = CustomDropdownMenu(widget=settings_button, corner_radius=0, bg_color=dropdown_bg_color, text_color=dropdown_text_color, hover_color=dropdown_hover_color)
appearance_sub_menu = settings_button_dropdown.add_submenu("Appearance")
appearance_sub_menu.add_option(option="Light", command=load_light_mode)
appearance_sub_menu.add_option(option="Dark", command=load_dark_mode)
appearance_sub_menu.add_option(option="System", command=load_system_mode)

# About Button Functions
def open_browser():
    webbrowser.open_new(SOURCE_CODE_URL)

about_button_dropdown = CustomDropdownMenu(widget=about_button, corner_radius=0, bg_color=dropdown_bg_color, text_color=dropdown_text_color, hover_color=dropdown_hover_color)
about_button_dropdown.add_option(option="Source Code", image=github_image, command=open_browser)

# TODO: Credits:
# Icons:
#   - Example : "Icon made by Pixel perfect from www.flaticon.com"
#   - Icon made by https://www.flaticon.com/authors/pixel-perfect from http://www.flaticon.com/
#   - Icon made by https://www.flaticon.com/authors/freepik from http://www.flaticon.com/

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
