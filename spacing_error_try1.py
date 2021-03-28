from lxml import etree
from bs4 import BeautifulSoup
import os
import tempfile
import shutil
import zipfile


def createNewDocx(origialDocx, newFilename):
    tmpDir = tempfile.mkdtemp()
    zip = zipfile.ZipFile(open(originalDocx, "rb"))
    zip.extractall(tmpDir)

    with zipfile.ZipFile("NDA-1-demo-one-page - date.docx", 'r') as zfp:
        with zfp.open('word/document.xml') as fp:
            soup = BeautifulSoup(fp.read(), 'xml')
            print()
    # get paragraphs
    for i, para in enumerate(soup.findAll("w:p")):
        for tags in para.findAll("w:t"):
            if tags.string == "Dattatraya":
                tags.string = tags.string.replace("Dattatraya", "Swapnil")
                print(tags.string)
                break

    with open(os.path.join(tmpDir, "word/document.xml"), "w", encoding="UTF-8") as f:
        f.write(xmlContent)
    filenames = zip.namelist()
    zipCopyFilename = newFilename
    with zipfile.ZipFile(zipCopyFilename, "w") as docx:
        for filename in filenames:
            docx.write(os.path.join(tmpDir, filename), filename)
    shutil.rmtree(tmpDir)


createNewDocx("NDA-1-demo-one-page - date.docx", "space_error.docx")
