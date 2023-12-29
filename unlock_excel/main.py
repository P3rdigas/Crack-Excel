import configparser
import os
import webbrowser
from tkinter import filedialog

import customtkinter
from CTkListbox import *
from CTkMenuBar import *
from CTkToolTip import *
from PIL import Image


class CrackExcel(customtkinter.CTk):
    APP_WIDTH = 800
    APP_HEIGHT = 600

    # Python colors: https://matplotlib.org/stable/gallery/color/named_colors.html or https://www.discogcodingacademy.com/turtle-colours
    MENUBAR_BACKGROUND_COLOR_LIGHT = "white"
    MENUBAR_BACKGROUND_COLOR_DARK = "black"

    DROPDOWN_BACKGROUND_COLOR_LIGHT = "white"
    DROPDOWN_BACKGROUND_COLOR_DARK = "grey20"

    DRAG_AND_DROP_BG_COLOR_LIGHT = "grey95"
    DRAG_AND_DROP_BG_COLOR_DARK = "grey15"
    DRAG_AND_DROP_SEPARATOR_COLOR_LIGHT = "grey90"
    DRAG_AND_DROP_SEPARATOR_COLOR_DARK = "black"

    TEXT_COLOR_LIGHT = "black"
    TEXT_COLOR_DARK = "white"
    HOVER_COLOR_LIGHT = "grey85"
    HOVER_COLOR_DARK = "grey25"

    GITHUB_LOGO_LIGHT_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "assets/icons/github-mark.png"
    )
    GITHUB_LOGO_DARK_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "assets/icons/github-mark-white.png"
    )
    ADD_FILE_LIGHT_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "assets/icons/add-file.png"
    )
    ADD_FILE_DARK_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "assets/icons/add-file-white.png"
    )
    DELETE_FILE_LIGHT_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "assets/icons/delete-file.png"
    )
    DELETE_FILE_DARK_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "assets/icons/delete-file-white.png"
    )

    SOURCE_CODE_URL = "https://github.com/P3rdigas/Crack-Excel"

    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Crack Excel")
        self.iconbitmap(
            os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "assets/logos/cracked_excel_logo_128x128.ico",
            )
        )
        self.geometry(f"{self.APP_WIDTH}x{self.APP_HEIGHT}")
        self.resizable(width=False, height=False)

        self.github_image = customtkinter.CTkImage(
            light_image=Image.open(self.GITHUB_LOGO_LIGHT_PATH),
            dark_image=Image.open(self.GITHUB_LOGO_DARK_PATH),
        )

        # Loading config file
        self.config_file = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "config.ini"
        )
        self.config = configparser.ConfigParser()
        self.load_configuration()

        # Set Menubar
        self.toolbar = CTkMenuBar(
            master=self,
            bg_color=(
                self.MENUBAR_BACKGROUND_COLOR_LIGHT,
                self.MENUBAR_BACKGROUND_COLOR_DARK,
            ),
        )
        self.file_button = self.toolbar.add_cascade(
            "File",
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        self.settings_button = self.toolbar.add_cascade(
            "Settings",
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        self.about_button = self.toolbar.add_cascade(
            "About",
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )

        self.file_button_dropdown = CustomDropdownMenu(
            widget=self.file_button,
            corner_radius=0,
            bg_color=(
                self.DROPDOWN_BACKGROUND_COLOR_LIGHT,
                self.DROPDOWN_BACKGROUND_COLOR_DARK,
            ),
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        self.file_button_dropdown.add_option(option="Import")
        self.file_button_dropdown.add_option(option="Export")
        self.file_button_dropdown.add_option(option="Exit", command=self.destroy)

        self.settings_button_dropdown = CustomDropdownMenu(
            widget=self.settings_button,
            corner_radius=0,
            bg_color=(
                self.DROPDOWN_BACKGROUND_COLOR_LIGHT,
                self.DROPDOWN_BACKGROUND_COLOR_DARK,
            ),
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        appearance_sub_menu = self.settings_button_dropdown.add_submenu("Appearance")
        appearance_sub_menu.add_option(
            option="Light", command=lambda: self.change_appearance_mode_event("light")
        )
        appearance_sub_menu.add_option(
            option="Dark", command=lambda: self.change_appearance_mode_event("dark")
        )
        appearance_sub_menu.add_option(
            option="System", command=lambda: self.change_appearance_mode_event("system")
        )

        self.about_button_dropdown = CustomDropdownMenu(
            widget=self.about_button,
            corner_radius=0,
            bg_color=(
                self.DROPDOWN_BACKGROUND_COLOR_LIGHT,
                self.DROPDOWN_BACKGROUND_COLOR_DARK,
            ),
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        self.about_button_dropdown.add_option(
            option="Source Code", image=self.github_image, command=self.open_browser
        )

        left_width = int(self.APP_WIDTH * 0.3)
        right_width = self.APP_WIDTH - left_width

        # Create Drag & Drop Frame
        drag_and_drop_frame = customtkinter.CTkFrame(
            self,
            corner_radius=0,
            width=left_width,
            fg_color=(
                self.DRAG_AND_DROP_BG_COLOR_LIGHT,
                self.DRAG_AND_DROP_BG_COLOR_DARK,
            ),
        )

        controls_frame = customtkinter.CTkFrame(
            drag_and_drop_frame, corner_radius=0, fg_color="transparent"
        )

        add_file_image = customtkinter.CTkImage(
            light_image=Image.open(self.ADD_FILE_LIGHT_PATH),
            dark_image=Image.open(self.ADD_FILE_DARK_PATH),
        )

        self.imported_files = set()

        add_file_button = customtkinter.CTkButton(
            controls_frame,
            image=add_file_image,
            text="",
            width=32,
            height=32,
            fg_color="transparent",
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
            command=self.open_file_explorer,
        )

        add_file_tooltip = CTkToolTip(
            add_file_button,
            message="Add file",
            corner_radius=5,
            border_width=1,
            border_color=("black", "white"),
        )

        add_file_button.bind(
            "<Enter>", lambda event: self.on_enter(event, add_file_tooltip)
        )

        delete_file_image = customtkinter.CTkImage(
            light_image=Image.open(self.DELETE_FILE_LIGHT_PATH),
            dark_image=Image.open(self.DELETE_FILE_DARK_PATH),
        )
        delete_file_button = customtkinter.CTkButton(
            controls_frame,
            image=delete_file_image,
            text="",
            width=32,
            height=32,
            fg_color="transparent",
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )

        delete_file_tooltip = CTkToolTip(
            delete_file_button,
            message="Delete file",
            corner_radius=5,
            border_width=1,
            border_color=("black", "white"),
        )

        delete_file_button.bind(
            "<Enter>", lambda event: self.on_enter(event, delete_file_tooltip)
        )

        controls_label = customtkinter.CTkLabel(controls_frame, text="Drag & Drop")

        # TODO: Add support for drag and drop
        drag_and_drop_separator = customtkinter.CTkFrame(
            drag_and_drop_frame,
            corner_radius=0,
            width=left_width,
            height=1,
            fg_color=(
                self.DRAG_AND_DROP_SEPARATOR_COLOR_LIGHT,
                self.DRAG_AND_DROP_SEPARATOR_COLOR_DARK,
            ),
            border_width=1,
        )

        self.files_listbox = CTkListbox(
            drag_and_drop_frame, border_width=0, justify="left"
        )

        # Create Separator Frame for Drag & Drop Frame and Execution Frame
        separator = customtkinter.CTkFrame(
            self,
            corner_radius=0,
            width=1,
            fg_color=(
                self.DRAG_AND_DROP_SEPARATOR_COLOR_LIGHT,
                self.DRAG_AND_DROP_SEPARATOR_COLOR_DARK,
            ),
            border_width=1,
        )

        # Create Execution Frame
        execution_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=right_width
        )

        # Drag & Drop Layout
        controls_label.pack(side="left", padx=10)
        delete_file_button.pack(side="right", padx=1)
        add_file_button.pack(side="right", padx=1)
        controls_frame.pack(fill="x")
        drag_and_drop_separator.pack(fill="x")
        self.files_listbox.pack(expand=True, fill="both")
        drag_and_drop_frame.pack(side="left", expand=True, fill="both")

        # Separator Layout
        separator.pack(side="left", fill="y")

        # Execution Frame Layout
        execution_frame.pack(side="left", expand=True, fill="both")

    def load_configuration(self):
        if os.path.isfile(self.config_file):
            self.config.read(self.config_file)
            theme = self.config.get("Settings", "theme")

            if theme == "system":
                customtkinter.set_appearance_mode("system")
            elif theme == "light":
                customtkinter.set_appearance_mode("light")
            else:
                customtkinter.set_appearance_mode("dark")
        else:
            self.config["Settings"] = {"theme": "system"}
            with open(self.config_file, "w") as configfile:
                self.config.write(configfile)

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

        self.config.set("Settings", "theme", new_appearance_mode)

        with open(self.config_file, "w") as configfile:
            self.config.write(configfile)

    def open_browser(self):
        webbrowser.open_new(self.SOURCE_CODE_URL)

    def on_enter(self, event, tooltip):
        tooltip.get()

    def open_file_explorer(self):
        files = filedialog.askopenfiles(filetypes=[("Excel files", "*.xlsx;*.xlsm")])

        if files:
            for file in files:
                file_path = file.name
                if file_path not in self.imported_files:
                    self.imported_files.add(file_path)
                    filename = os.path.basename(file_path)
                    self.files_listbox.insert("END", filename, text_anchor="w")


def main():
    app = CrackExcel()

    # TODO #1: Credits:
    # Icons:
    #   - Example : "Icon made by Pixel perfect from www.flaticon.com"
    #   - Icon made by https://www.flaticon.com/authors/pixel-perfect from http://www.flaticon.com/
    #   - Icon made by https://www.flaticon.com/authors/freepik from http://www.flaticon.com/
    #   - Icon made by https://www.flaticon.com/authors/ivan-abirawa from http://www.flaticon.com/
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
