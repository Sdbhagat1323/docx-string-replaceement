import zipfile
from lxml import etree
from bs4 import BeautifulSoup
import os
import tempfile
import shutil
import re

# str(input("Enter the string to be replace here.."))
INPUT_TEXT = "Effective Date"
NEW_STRING = "Date"  # str(input("Enter the string to be replace here.."))


def createNewDocx(originalDocx, newFilename):
    # make temdir from original doc
    # temDir = Docx
    tmpDir = tempfile.mkdtemp()
    zip = zipfile.ZipFile(open(originalDocx, "rb"))
    zip.extractall(tmpDir)

    with zipfile.ZipFile(originalDocx, 'r') as zfp:
        with zfp.open(zfp, 'r') as fp:
            soup = BeautifulSoup(fp.read(), 'html_parser')
            print(soup.prettify())
            # parsing with tags

            for i, para in enumerate(soup.findAll("w:p")):
                for tag in soup.findAll("w:t"):
                    if tag.string == INPUT_TEXT:
                        print("string Found")
                        tag.string = tag.string.replace(
                            INPUT_TEXT, NEW_STRING)
                        print(tag.string)
                


    print("--------------------------------------------------------------------------------------------")
    xml_string = soup.decode('UTF-8')
    print(xml_string)

    # Use temdir to write updated xml file.
    # make sure to decode in "UTF-8"  or file crash
    # check for xmlContent format (byte or string)

    with open(os.path.join(tmpDir, "word/document.html"), "wb", encoding="UTF-8") as f:
        f.write(xml_string)

    # get all file name from zip object and copy them to zipcopyFilename
    filenames = zip.namelist()
    zipCopyFilename = newFilename

    # write doc from filenames list
    with zipfile.ZipFile(zipCopyFilename, "w") as docx:
        for filename in filenames:
            docx.write(os.path.join(tmpDir, filename), filename)
    shutil.rmtree(tmpDir)


createNewDocx("NDA-1-demo-one-page - date.docx", "libra_docx_file.docx")
