import zipfile
from lxml import etree
from bs4 import BeautifulSoup
import os
import tempfile
import shutil


def createNewDocx(originalDocx,xmlContent,newFilename):
    tmpDir = tempfile.mkdtemp()
    zip = zipfile.ZipFile(open(originalDocx,"rb"))
    zip.extractall(tmpDir)
    with open(os.path.join(tmpDir,"word/document.xml"), "w", encoding="UTF-8") as f:
        f.write(xmlContent)
    filenames = zip.namelist()
    zipCopyFilename = newFilename
    with zipfile.ZipFile(zipCopyFilename,"w") as docx:
        for filename in filenames:
            docx.write(os.path.join(tmpDir,filename),filename)
    shutil.rmtree(tmpDir)



with zipfile.ZipFile("NDA-1-demo-one-page - date.docx", 'r') as zfp:
    with zfp.open('word/document.xml') as fp:
        soup = BeautifulSoup(fp.read(), 'xml')
print(soup.prettify())
'''

paragraphs = []
texts = []
for node in tree.iter(tag=etree.Element):
    texts.append(node.text)
    

    if node.text == "VirtualWeb":
        node1 = node.text.replace("VirtualWeb",'WoodWork')
            
        print("replacement done")
        print(node1)
print(texts)
'''

for tag in soup.findAll("w:t"):
    if tag.string == "Dattatraya":
        print("-----------------------")
        print(tag.string)
        tag.string = tag.string.replace('Dattatraya', ' Swapnil ')
        print("-----------------------")
        print(tag.string)
        break

    

xml_string = soup.decode('utf-8')


createNewDocx("NDA-1-demo-one-page - date.docx", xml_string, "fun_output.docx")




#Extra code 
'''

f = open("beautifull_soup4.html", "w")
f.write(str(soup))
f.close()

with zipfile.ZipFile("beautifull_soup4.html", "w") as html:
            for filename in filenames:
                html.write(os.path.join(tmp_dir,filename), filename)
'''