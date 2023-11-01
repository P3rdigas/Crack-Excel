from io import BytesIO
from lxml import etree
from zipfile import ZipFile

CONST_DATA_ZIP = "data.zip"
CONST_ZIP_EXTENSION = ".zip"

CONST_XL_FOLDER = "xl/"
CONST_WORKSHEETS_FOLDER = "worksheets/"

CONST_WORKBOOK_FILE = "workbook.xml"

CONST_READONLY_PROTECTION = "fileSharing"
CONST_WORKBOOK_PROTECTION = "workbookProtection"

CONST_UTF = "utf-8"

def convert_excel_to_zip(excel_file):
    with open(excel_file, "rb") as f:
        excel_data = f.read()

    excel_to_zip = BytesIO()

    # Store the bytes of the excel inside of an zip file
    with ZipFile(excel_to_zip, "w") as zf:
        zf.writestr(CONST_DATA_ZIP, excel_data)

    zip_file = BytesIO()

    # Extracts the zip file (data.zip) content to an in memory variable
    with ZipFile(excel_to_zip, 'r') as zf:
        file_content = zf.read(CONST_DATA_ZIP)
        zip_file.write(file_content)

    excel_to_zip.close()

    return zip_file

def remove_protections(xml_content, tag):
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
    modified_xml_content = etree.tostring(root, encoding=CONST_UTF)

    return modified_xml_content
   
def crack_excel(zip_file):
    with ZipFile(zip_file, "r") as zf:
        # Removes the workbook protections
        workbook_file = CONST_XL_FOLDER + CONST_WORKBOOK_FILE

        if workbook_file in zf.namelist():
            with zf.open(workbook_file, "r") as workbook:
                workbook_content_string = workbook.read()       

            modified_workbook_content = remove_protections(workbook_content_string, CONST_WORKBOOK_PROTECTION)
            modified_workbook_content = remove_protections(modified_workbook_content, CONST_READONLY_PROTECTION)

            cracked_zip_file = BytesIO()

            # Save the modified content to a new zip file in memory
            with ZipFile(cracked_zip_file, "w") as czf:
                for item in zf.infolist():
                    if item.filename == workbook_file:
                        czf.writestr(item.filename, modified_workbook_content)
                    else:
                        content = zf.read(item.filename)
                        czf.writestr(item.filename, content)
        
        # TODO: Unprotect sheets
    
    zip_file.close()

    return cracked_zip_file

def create_unprotected_file(zip_file, filename, extension):
    file_path = filename + " - CRACKED" + extension

    with open(file_path, 'wb') as excel_file:
        excel_file.write(zip_file.getvalue())

    print(f"Excel file saved as {file_path}")