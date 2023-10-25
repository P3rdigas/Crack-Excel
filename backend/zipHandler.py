import io
import zipfile

CONST_DATA_ZIP = "data.zip"
CONST_ZIP_EXTENSION = ".zip"

CONST_XL_FOLDER = "xl/"
CONST_WORKSHEETS_FOLDER = "worksheets/"

CONST_WORKBOOK_FILE = "workbook.xml"

CONST_WORKBOOK_PROTECTION = "<fileSharing"

# Converts an Excel file to a zip file TODO: in memory
# Args:
#   excel_file: The filename with the extension of the Excel file.
# Returns:
#   TODO: A ZipFile object containing the Excel file.
def convert_excel_to_zip(excel_file):
    # Reads the bytes of the excel file
    with open(excel_file, "rb") as f:
        excel_data = f.read()

    # Create an in-memory buffer for the creating the first zip
    excel_to_zip_buffer = io.BytesIO()

    # Creates the zip file that will contain the Excel data
    with zipfile.ZipFile(excel_to_zip_buffer, "w") as zf:
        zf.writestr(CONST_DATA_ZIP, excel_data)

    # TODO: Possible optimization, extract the content of the zf zip and he will get the data.zip
    # Open the first zip file in append mode
    with zipfile.ZipFile(excel_to_zip_buffer, "a") as converted_zip:
        if CONST_DATA_ZIP in converted_zip.namelist():
            # Extract the data zip file
            data_zip = converted_zip.read(CONST_DATA_ZIP)

            # Create a new in-memory buffer for the new zip file
            zip_file_buffer = io.BytesIO()

            # Open the data zip file from memory
            with zipfile.ZipFile(io.BytesIO(data_zip), 'r') as data:
                # Extract the content from the inner zip file
                for file_name in data.namelist():
                    data_content = data.read(file_name)
                    # Write the content to the root of the converted_zip
                    converted_zip.writestr(file_name, data_content)


            # Create a new ZipFile
            with zipfile.ZipFile(zip_file_buffer, "w") as new_zip_file:
            
                # Copy the files you want to keep to the new ZipFile
                for item in converted_zip.infolist():
                    if item.filename != CONST_DATA_ZIP:
                        data = converted_zip.read(item.filename)
                        new_zip_file.writestr(item, data)

    excel_to_zip_buffer.close()

    return zip_file_buffer

def crack_excel(zip_file):
    with zipfile.ZipFile(zip_file, "r") as zf:
        workbook_file = CONST_XL_FOLDER + CONST_WORKBOOK_FILE

        if workbook_file in zf.namelist():
            with zf.open(workbook_file) as workbook:
                if CONST_WORKBOOK_PROTECTION in workbook:
                    print("O ficheiro está protegido")
                else:
                    print("O ficheiro não está protegido")

        worksheets_folder = CONST_XL_FOLDER + CONST_WORKSHEETS_FOLDER
        