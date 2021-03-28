import zipfile
from lxml import etree
from bs4 import BeautifulSoup
import os
import tempfile
import shutil



def createNewDocx(originalDocx,newFilename):
    tmpDir = tempfile.mkdtemp()
    zip = zipfile.ZipFile(open(originalDocx,"rb"))
    zip.extractall(tmpDir)

    with open(os.path.join(tmpDir,"word/document.xml"),"w") as f:
        soup = BeautifulSoup(xmlContent.read(), 'xml')
        print(soup.prettify())
    
    
    for tag in soup.findAll("w:t"):
        if tag.string == "Confidentiality Agreement":
            print("-----------------------")
            print(tag.string)
            tag.string = tag.string.replace('Confidentiality Agreement', 'Rent Agreement')
            print("-----------------------")
            print(tag.string)
            break

        else:
            print("NOthing is match..")

        print(soup.prettify())
    


    filenames = zip.namelist(soup)
    zipCopyFilename = newFilename
    with zipfile.ZipFile(zipCopyFilename,"w") as docx:
        for filename in filenames:
            docx.write(os.path.join(tmpDir,filename),filename)
    shutil.rmtree(tmpDir)


createNewDocx("NDA-1-demo-one-page - date.docx", "fun_docx_conver.docx")