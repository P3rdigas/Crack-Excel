import io
import zipfile

CONST_ZIP_EXTENSION = ".zip"
CONST_DATA_ZIP = "data.zip"

# Converts an Excel file to a zip file TODO: in memory
# Args:
#   excel_file: The filename with the extension of the Excel file.
#   excel_name: The name of the Excel file
# Returns:
#   TODO: A ZipFile object containing the Excel file.
def convert_excel_to_zip(excel_file, excel_name):
    # Reads the bytes of the excel file
    with open(excel_file, "rb") as f:
        excel_data = f.read()

    # Create an in-memory buffer for the creating the first zip
    excel_to_zip_buffer = io.BytesIO()
    
    zip_file_name = excel_name + CONST_ZIP_EXTENSION

    # Creates the zip file that will contain the Excel data
    with zipfile.ZipFile(excel_to_zip_buffer, "w") as zf:
        zf.writestr(CONST_DATA_ZIP, excel_data)

    # TODO: Possible optimization, extract the content of the zf zip and he will get the data.zip
    # Open the first zip file in append mode
    with zipfile.ZipFile(excel_to_zip_buffer, "a") as converted_zip:
        if CONST_DATA_ZIP in converted_zip.namelist():
            # Extract the data zip file
            data_zip = converted_zip.read(CONST_DATA_ZIP)

            # Open the data zip file from memory
            with zipfile.ZipFile(io.BytesIO(data_zip), 'r') as data:
                # Extract the content from the inner zip file
                for file_name in data.namelist():
                    data_content = data.read(file_name)
                    # Write the content to the root of the converted_zip
                    converted_zip.writestr(file_name, data_content)


            # Create a new ZipFile
            with zipfile.ZipFile(zip_file_name, "w") as new_zip_file:
            
                # Copy the files you want to keep to the new ZipFile
                for item in converted_zip.infolist():
                    if item.filename != CONST_DATA_ZIP:
                        data = converted_zip.read(item.filename)
                        new_zip_file.writestr(item, data)

    excel_to_zip_buffer.close()




# #
# # Open the zip file 
# def open_zip_file():
#     print("ADeus")

# #
# # Function that converts the excel file to a zip and creates the zip file 
# def convert_to_zip(excel_name, excel_file_path, source_folder_path):
#     zip_file_name = excel_name + CONST_ZIP_EXTENSION
#     zip_file_path = os.path.join(source_folder_path, zip_file_name)
#     os.rename(excel_file_path, zip_file_path)