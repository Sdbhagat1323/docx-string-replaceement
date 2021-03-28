from lxml import etree
import zipfile
import xml.etree.ElementTree as ET 

# open word file in xml format

with open("test_template.docx", "r+b") as f:
    zip = zipfile.ZipFile(f)
    xml_content = zip.read("word/document.xml")
    #print(xml_content)
#Now we are getting the file out xml

#get etree from xml
#To get byte data "tree"
tree = etree.fromstring(xml_content)
print(type(tree))
# To get string data "tree_2"
#tree_2 = etree.tostring(tree, pretty_print=True)
#print(type(tree_2))

#Now getting all the parsing data
# getting trought every node of the text

#for node in tree.iter(tag=etree.Element):
         #print(node, node.text)

TO_REPLACE  = "VirtualWeb"
REPLACE_WITH = "Wonder Woodwork"


paragraphs = []
texts = []
for node in tree.iter(tag=etree.Element):
    texts.append(node.text)
    

    if node.text == "VirtualWeb":
        node1 = node.text.replace("VirtualWeb",'WoodWork')
            
        print("replacement done")
        print(node1)
print(texts)









tree2 = etree.tostring(tree)
tree2.write("output.xml",  xml_declaration=True, method='xml', encoding="utf8")
#print(etree.tostring(tree, pretty_print=True))
'''

def _write_and_close_docx (self, xml_content, output_filename):
        """ Create a temp directory, expand the original docx zip.
            Write the modified xml to word/document.xml
            Zip it up as the new docx
        """

        tmp_dir = tempfile.mkdtemp()

        self.zipfile.extractall(tmp_dir)

        with open(os.path.join(tmp_dir,'word/document.xml'), 'w') as f:
            xmlstr = etree.tostring (xml_content, pretty_print=True)
            f.write(xmlstr)

        # Get a list of all the files in the original docx zipfile
        filenames = self.zipfile.namelist()
        # Now, create the new zip file and add all the filex into the archive
        zip_copy_filename = output_filename
        with zipfile.ZipFile(zip_copy_filename, "w") as docx:
            for filename in filenames:
                docx.write(os.path.join(tmp_dir,filename), filename)

        # Clean up the temp dir
        shutil.rmtree(tmp_dir)



def _join_tags(self, my_etree):
        chars = []
        openbrac = False
        inside_openbrac_node = False

        for node,text in self._itertext(my_etree):
            # Scan through every node with text
            for i,c in enumerate(text):
                # Go through each node's text character by character
                if c == '[':
                    openbrac = True # Within a tag
                    inside_openbrac_node = True # Tag was opened in this node
                    openbrac_node = node # Save ptr to open bracket containing node
                    chars = []
                elif c== ']':
                    assert openbrac
                    if inside_openbrac_node:
                        # Open and close inside same node, no need to do anything
                        pass
                    else:
                        # Open bracket in earlier node, now it's closed
                        # So append all the chars we've encountered since the openbrac_node '['
                        # to the openbrac_node
                        chars.append(']')
                        openbrac_node.text += ''.join(chars)
                        # Also, don't forget to remove the characters seen so far from current node
                        node.text = text[i+1:]
                    openbrac = False
                    inside_openbrac_node = False
                else:
                    # Normal text character
                    if openbrac and inside_openbrac_node:
                        # No need to copy text
                        pass
                    elif openbrac and not inside_openbrac_node:
                        chars.append(c)
                    else:
                        # outside of a open/close
                    pass
           if openbrac and not inside_openbrac_node:
                # Went through all text that is part of an open bracket/close bracket
                # in other nodes
                # need to remove this text completely
                node.text = ""
            inside_openbrac_node = False

'''