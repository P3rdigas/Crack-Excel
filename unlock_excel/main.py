import os
import webbrowser
import configparser
import customtkinter
from PIL import Image
from CTkMenuBar import *

class CrackExcel(customtkinter.CTk):
    # Python colors: https://matplotlib.org/stable/gallery/color/named_colors.html or https://www.discogcodingacademy.com/turtle-colours
    MENUBAR_BACKGROUND_COLOR_LIGHT = "white"
    MENUBAR_BACKGROUND_COLOR_DARK = "black"

    DROPDOWN_BACKGROUND_COLOR_LIGHT = "white"
    DROPDOWN_BACKGROUND_COLOR_DARK = "grey20"

    TEXT_COLOR_LIGHT = "black"
    TEXT_COLOR_DARK = "white"
    HOVER_COLOR_LIGHT = "light gray"
    HOVER_COLOR_DARK = "grey25"

    GITHUB_LOGO_LIGHT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets/icons/github-mark.png')
    GITHUB_LOGO_DARK_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets/icons/github-mark-white.png')

    SOURCE_CODE_URL = "https://github.com/P3rdigas/Crack-Excel"

    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Crack Excel")
        self.iconbitmap(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets/logos/cracked_excel_logo_128x128.ico'))
        self.geometry(f"{800}x{600}")
        self.resizable(width=False, height=False)

        # https://github.com/Akascape/CTkMenuBar
        self.menu_bar_bg = None
        self.menu_bar_text_color = None
        self.menu_bar_hover_color = None
        self.dropdown_bg_color = None
        self.dropdown_text_color = None
        self.dropdown_hover_color = None
        self.github_image = customtkinter.CTkImage(light_image=Image.open(self.GITHUB_LOGO_LIGHT_PATH), dark_image=Image.open(self.GITHUB_LOGO_DARK_PATH))

        # Loading config file
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.ini')
        self.config = configparser.ConfigParser()
        self.load_configuration()        

        self.toolbar = CTkMenuBar(master=self, bg_color=self.menu_bar_bg)
        self.file_button = self.toolbar.add_cascade("File", text_color=self.menu_bar_text_color, hover_color=self.menu_bar_hover_color)
        self.settings_button = self.toolbar.add_cascade("Settings", text_color=self.menu_bar_text_color, hover_color=self.menu_bar_hover_color)
        self.about_button = self.toolbar.add_cascade("About", text_color=self.menu_bar_text_color, hover_color=self.menu_bar_hover_color)

        self.file_button_dropdown = CustomDropdownMenu(widget=self.file_button, corner_radius=0, bg_color=self.dropdown_bg_color, text_color=self.dropdown_text_color, hover_color=self.dropdown_hover_color)
        self.file_button_dropdown.add_option(option="Import")
        self.file_button_dropdown.add_option(option="Export")
        self.file_button_dropdown.add_option(option="Exit", command=self.destroy)

        self.settings_button_dropdown = CustomDropdownMenu(widget=self.settings_button, corner_radius=0, bg_color=self.dropdown_bg_color, text_color=self.dropdown_text_color, hover_color=self.dropdown_hover_color)
        appearance_sub_menu = self.settings_button_dropdown.add_submenu("Appearance")
        appearance_sub_menu.add_option(option="Light", command=self.load_light_mode)
        appearance_sub_menu.add_option(option="Dark", command=self.load_dark_mode)
        appearance_sub_menu.add_option(option="System", command=self.load_system_mode)

        self.about_button_dropdown = CustomDropdownMenu(widget=self.about_button, corner_radius=0, bg_color=self.dropdown_bg_color, text_color=self.dropdown_text_color, hover_color=self.dropdown_hover_color)
        self.about_button_dropdown.add_option(option="Source Code", image=self.github_image, command=self.open_browser)

    def load_configuration(self):
        if os.path.isfile(self.config_file):
            self.config.read(self.config_file)
            theme = self.config.get('Settings', 'theme')

            if theme == "system":
                self.load_init_system_mode()
            elif theme == "light":
                customtkinter.set_appearance_mode("light")
                self.load_init_light_mode()
            else:
                customtkinter.set_appearance_mode("dark")
                self.load_init_dark_mode()
        else:
            self.config['Settings'] = {'theme': 'system'}
            with open(self.config_file, 'w') as configfile:
                self.config.write(configfile)
            
            self.load_init_system_mode()
    
    def load_init_light_mode(self):
        self.menu_bar_bg = self.MENUBAR_BACKGROUND_COLOR_LIGHT
        self.menu_bar_text_color = self.TEXT_COLOR_LIGHT
        self.menu_bar_hover_color = self.HOVER_COLOR_LIGHT
        self.dropdown_bg_color = self.DROPDOWN_BACKGROUND_COLOR_LIGHT
        self.dropdown_text_color = self.TEXT_COLOR_LIGHT
        self.dropdown_hover_color = self.HOVER_COLOR_LIGHT

    def load_init_dark_mode(self):
        self.menu_bar_bg = self.MENUBAR_BACKGROUND_COLOR_DARK
        self.menu_bar_text_color = self.TEXT_COLOR_DARK
        self.menu_bar_hover_color = self.HOVER_COLOR_DARK
        self.dropdown_bg_color = self.DROPDOWN_BACKGROUND_COLOR_DARK
        self.dropdown_text_color = self.TEXT_COLOR_DARK
        self.dropdown_hover_color = self.HOVER_COLOR_DARK

    def load_init_system_mode(self):
        customtkinter.set_appearance_mode("system")
        appearance_mode = customtkinter.get_appearance_mode()

        if appearance_mode == "light":
            self.load_init_light_mode()
        else:
            self.load_init_dark_mode()

    def load_images(self, text):
        match text:
            case "Source Code":
                return customtkinter.CTkImage(light_image=Image.open(self.GITHUB_LOGO_LIGHT_PATH), dark_image=Image.open(self.GITHUB_LOGO_DARK_PATH))

    def updated_dropdown(self, widget, original_dropdown, bg_color, text_color, hover_color):
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
                        image = self.load_images(option.cget("text"))
                        
                    submenu_copy.add_option(
                        option=sub_option.cget("text"),
                        image = image,
                        command=sub_option._command
                    )
            else:
                image = None
                
                if hasattr(option, '_image') and option._image is not None:
                    image = self.load_images(option.cget("text"))

                new_dropdown.add_option(
                    option=option.cget("text"),
                    image = image,
                    command=option._command
                )

        return new_dropdown

    def configure_light_mode(self):
        customtkinter.set_appearance_mode("light")

        self.toolbar.configure(bg_color=self.MENUBAR_BACKGROUND_COLOR_LIGHT)
        self.file_button.configure(text_color=self.TEXT_COLOR_LIGHT, hover_color=self.HOVER_COLOR_LIGHT)
        self.settings_button.configure(text_color=self.TEXT_COLOR_LIGHT, hover_color=self.HOVER_COLOR_LIGHT)
        self.about_button.configure(text_color=self.TEXT_COLOR_LIGHT, hover_color=self.HOVER_COLOR_LIGHT)

        self.file_button_dropdown = self.updated_dropdown(self.file_button, self.file_button_dropdown, self.DROPDOWN_BACKGROUND_COLOR_LIGHT, self.TEXT_COLOR_LIGHT, self.HOVER_COLOR_LIGHT)
        self.settings_button_dropdown = self.updated_dropdown(self.settings_button, self.settings_button_dropdown, self.DROPDOWN_BACKGROUND_COLOR_LIGHT, self.TEXT_COLOR_LIGHT, self.HOVER_COLOR_LIGHT)
        self.about_button_dropdown = self.updated_dropdown(self.about_button, self.about_button_dropdown, self.DROPDOWN_BACKGROUND_COLOR_LIGHT, self.TEXT_COLOR_LIGHT, self.HOVER_COLOR_LIGHT)

    def load_light_mode(self):
        self.configure_light_mode()
        
        self.config.set('Settings', 'theme', 'light')

        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def configure_dark_mode(self):        
        customtkinter.set_appearance_mode("dark")

        self.toolbar.configure(bg_color=self.MENUBAR_BACKGROUND_COLOR_DARK)
        self.file_button.configure(text_color=self.TEXT_COLOR_DARK, hover_color=self.HOVER_COLOR_DARK)
        self.settings_button.configure(text_color=self.TEXT_COLOR_DARK, hover_color=self.HOVER_COLOR_DARK)
        self.about_button.configure(text_color=self.TEXT_COLOR_DARK, hover_color=self.HOVER_COLOR_DARK)

        self.file_button_dropdown = self.updated_dropdown(self.file_button, self.file_button_dropdown, self.DROPDOWN_BACKGROUND_COLOR_DARK, self.TEXT_COLOR_DARK, self.HOVER_COLOR_DARK)
        self.settings_button_dropdown = self.updated_dropdown(self.settings_button, self.settings_button_dropdown, self.DROPDOWN_BACKGROUND_COLOR_DARK, self.TEXT_COLOR_DARK, self.HOVER_COLOR_DARK)
        self.about_button_dropdown = self.updated_dropdown(self.about_button, self.about_button_dropdown, self.DROPDOWN_BACKGROUND_COLOR_DARK, self.TEXT_COLOR_DARK, self.HOVER_COLOR_DARK)

    def load_dark_mode(self):
        self.configure_dark_mode()

        self.config.set('Settings', 'theme', 'dark')

        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)
    
    def load_system_mode(self):
        customtkinter.set_appearance_mode("system")
        appearance_mode = customtkinter.get_appearance_mode()
    
        if appearance_mode == "light":
            self.configure_light_mode()
        else:
            self.configure_dark_mode()
    
        self.config.set('Settings', 'theme', 'system')
    
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def open_browser(self):
        webbrowser.open_new(self.SOURCE_CODE_URL)

def main():
    app = CrackExcel()

    # TODO #1: Credits:
    # Icons:
    #   - Example : "Icon made by Pixel perfect from www.flaticon.com"
    #   - Icon made by https://www.flaticon.com/authors/pixel-perfect from http://www.flaticon.com/
    #   - Icon made by https://www.flaticon.com/authors/freepik from http://www.flaticon.com/
    app.mainloop()

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

if __name__ == "__main__":
    main()