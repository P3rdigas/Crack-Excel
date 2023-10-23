import zipfile

CONST_ZIP_EXTENSION = ".zip"


#
# Open the zip file 
def open_zip_file():
    print("ADeus")

#
# Function that converts the excel file to a zip and creates the zip file 
def convert_to_zip(excel_name, excel_file_path, source_folder_path):
    zip_file_name = excel_name + CONST_ZIP_EXTENSION
    zip_file_path = os.path.join(source_folder_path, zip_file_name)
    os.rename(excel_file_path, zip_file_path)