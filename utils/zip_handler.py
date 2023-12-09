import tempfile
from io import BytesIO
from lxml import etree
from zipfile import ZipFile

CONST_DATA_ZIP = "data.zip"
CONST_ZIP_EXTENSION = ".zip"
CONST_XML_EXTENSION = ".xml"

CONST_XL_FOLDER = "xl/"
CONST_WORKSHEETS_FOLDER = "worksheets/"

CONST_SHEET_FILE = "sheet"
CONST_VBA_FILE = "vbaProject.bin"
CONST_WORKBOOK_FILE = "workbook.xml"

CONST_READONLY_PROTECTION = "fileSharing"
CONST_SHEET_PROTECTION = "sheetProtection"
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

def remove_protection(xml_content, tag):
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
        vba_file = CONST_XL_FOLDER + CONST_VBA_FILE
        workbook_file = CONST_XL_FOLDER + CONST_WORKBOOK_FILE
        sheets_files = CONST_XL_FOLDER + CONST_WORKSHEETS_FOLDER + CONST_SHEET_FILE

        cracked_zip_file = BytesIO()

        modified_file_content = None

        for file in zf.namelist():
            if file == vba_file:
                with zf.open(file, "r") as vba:
                    vba_content = vba.read()
                
                modified_file_content = vba_content.decode("utf-8").replace("DPB", "DPX").encode("utf-8")
                #print(modified_file_content)

            elif file == workbook_file:
                with zf.open(file, "r") as workbook:
                    workbook_content = workbook.read()       

                modified_file_content = remove_protection(workbook_content, CONST_WORKBOOK_PROTECTION)
                modified_file_content = remove_protection(modified_file_content, CONST_READONLY_PROTECTION)

            elif file.startswith(sheets_files) and file.endswith(CONST_XML_EXTENSION) and file != sheets_files:
                with zf.open(file, "r") as sheet:
                    sheet_content = sheet.read()
                
                modified_file_content = remove_protection(sheet_content, CONST_SHEET_PROTECTION)

            # Save the modified content to a new zip file in memory
            with ZipFile(cracked_zip_file, "a") as czf:
                if modified_file_content is None:
                    content = zf.read(file)
                    czf.writestr(file, content)
                else:
                    czf.writestr(file, modified_file_content)

            modified_file_content = None

    zip_file.close()

    return cracked_zip_file

def change_vba(zip_file):
    # Create a temporary file and write the VBA bytes to it
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as temp_file:
        temp_file.write(zip_file.getvalue())

def create_unprotected_file(zip_file, filename, extension):
    file_path = filename + " - CRACKED" + extension

    with open(file_path, 'wb') as excel_file:
        excel_file.write(zip_file.getvalue())

    print(f"Excel file saved as {file_path}")