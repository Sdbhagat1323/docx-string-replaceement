from lxml import etree
import zipfile
import xml.etree.ElementTree as ET 

# open word file in xml format

with open("test_template.docx", "rb") as f:
    zip = zipfile.ZipFile(f)
    xml_content = zip.read("word/document.xml")
    #print(xml_content)

tree = ET.parse(xml_content)
root = tree.getroot()
'''
for elem in root.getiterator():
    elem.text = elem.text.replace('VirtualWeb', 'WoodWork')
    elem.text = elem.text.replace('Dattatraya', 'Swapnil')
root.write("output.xml",  xml_declaration=True, method='xml', encoding="utf8")
'''


root.find("VirtualWeb").text = tree.find("VirtualWeb").text.replace("VirtualWeb", 'WoodWork')

print(root)