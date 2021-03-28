#import mammoth 
import docx
import os
#import pypandoc
'''

doc = docx.Document("NDA-1-demo-one-page - date.docx")

with open("NDA-1-demo-one-page - date.docx", "rb") as f:
    result = mammoth.convert_to_html(f)
    html = result.value
print(type(html))
print(html)


print(type(result))
'''
import os
path =  os.path.abspath("NDA-1-demo-one-page - date.docx")


import win32com.client

doc = win32com.client.GetObject(path)
doc.SaveAs(FileName="demo.html", FileFormat=8)
doc.Close ()


'''
output = pypandoc.convert(source='demo.html', format='html', to='docx', outputfile='output.docx', extra_args=['-RTS'])


output = pypandoc.convert_file('demo.html', 'docx', outputfile="pypandocx_file.docx")
assert output == ""
'''