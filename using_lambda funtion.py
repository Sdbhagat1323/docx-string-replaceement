import re

from lxml import etree
import zipfile
import xml.etree.ElementTree as ET 



with open("test_template.docx", "r+b") as f:
    zip = zipfile.ZipFile(f)
    xml_content = zip.read("word/document.xml")
    tree = etree.fromstring(xml_content)
    tree_2 = etree.tostring(tree, pretty_print=True)

substitutions = {'VirtualWeb': 'myvalue'}
pattern = re.compile(r'%([^%]+)%')
tree = re.sub(pattern, lambda m: substitutions[m.group(1)], tree)

print(tree)



'''
#original
xmlstring = open('myxmldocument.xml', 'r').read()
substitutions = {'SITEDESCR': 'myvalue', ...}
pattern = re.compile(r'%([^%]+)%')
xmlstring = re.sub(pattern, lambda m: substitutions[m.group(1)], xmlstring)

'''