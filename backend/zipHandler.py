from io import BytesIO
from lxml import etree
from zipfile import ZipFile


CONST_DATA_ZIP = "data.zip"
CONST_ZIP_EXTENSION = ".zip"

CONST_XL_FOLDER = "xl/"
CONST_WORKSHEETS_FOLDER = "worksheets/"

CONST_WORKBOOK_FILE = "workbook.xml"

CONST_WORKBOOK_PROTECTION = "workbookProtection"

CONST_UTF = "utf-8"

def convert_excel_to_zip(excel_file):
    with open(excel_file, "rb") as f:
        excel_data = f.read()

    excel_to_zip_buffer = BytesIO()

    with ZipFile(excel_to_zip_buffer, "w") as zf:
        zf.writestr(CONST_DATA_ZIP, excel_data)

    # TODO: Possible optimization, extract the content of the zf zip and he will get the data.zip
    # Open the first zip file in append mode
    with ZipFile(excel_to_zip_buffer, "a") as converted_zip:
        if CONST_DATA_ZIP in converted_zip.namelist():
            # Extract the data zip file
            data_zip = converted_zip.read(CONST_DATA_ZIP)

            # Create a new in-memory buffer for the new zip file
            zip_file_buffer = BytesIO()

            # Open the data zip file from memory
            with ZipFile(BytesIO(data_zip), 'r') as data:
                # Extract the content from the inner zip file
                for file_name in data.namelist():
                    data_content = data.read(file_name)
                    # Write the content to the root of the converted_zip
                    converted_zip.writestr(file_name, data_content)


            # Create a new ZipFile
            with ZipFile(zip_file_buffer, "w") as new_zip_file:
            
                # Copy the files you want to keep to the new ZipFile
                for item in converted_zip.infolist():
                    if item.filename != CONST_DATA_ZIP:
                        data = converted_zip.read(item.filename)
                        new_zip_file.writestr(item, data)

    excel_to_zip_buffer.close()

    return zip_file_buffer

def remove_workbook_protection(xml_content, tag):
    tree = etree.parse(BytesIO(xml_content))
    root = tree.getroot()
    
    elements_to_remove = root.xpath(".//*[local-name() = '{}']".format(tag))

    if elements_to_remove:
        for element in elements_to_remove:
            parent = element.getparent()
            parent.remove(element)
        print("{} removed".format(tag))
    else:
        print("{} doesn't exists in the file".format(tag))

    # Serialize the modified XML content to a string
    modified_xml_content = etree.tostring(root, encoding=CONST_UTF).decode(CONST_UTF)

    return modified_xml_content
   
def crack_excel(zip_file):
    with ZipFile(zip_file, "r") as zf:
        # Removes the workbook protection
        workbook_file = CONST_XL_FOLDER + CONST_WORKBOOK_FILE

        if workbook_file in zf.namelist():
            with zf.open(workbook_file, "r") as workbook:
                workbook_content_string = workbook.read()       

            modified_workbook_content = remove_workbook_protection(workbook_content_string, CONST_WORKBOOK_PROTECTION)

            cracked_zip_file = BytesIO()

            # Save the modified content to a new zip file in memory
            with ZipFile(cracked_zip_file, "w") as czf:
                for item in zf.infolist():
                    if item.filename == workbook_file:
                        czf.writestr(item.filename, modified_workbook_content)
                    else:
                        content = zf.read(item.filename)
                        czf.writestr(item.filename, content) 

    zip_file.close()

    
    # # Creates a zip file on disk
    # with ZipFile(cracked_zip_file, "r") as czf:
    #     # Save the modified content to a new zip file on disk
    #     with ZipFile("output.zip", "w") as output_zip:
    #         for item in czf.infolist():
    #             content = czf.read(item.filename)
    #             output_zip.writestr(item.filename, content)

    return cracked_zip_file

def create_unprotected_file(zip_file, filename, extension):
    file_path = filename + " - CRACKED" + extension

    with open(file_path, 'wb') as excel_file:
        excel_file.write(zip_file.getvalue())

    print(f"Excel file saved as {file_path}")