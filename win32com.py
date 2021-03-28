
import os
import win32com.client 
path =  os.path.abspath("NDA-1-demo-one-page - date.docx")



doc = win32com.client.GetObject (path)
doc.SaveAs (FileName="test.html", FileFormat=8)
doc.Close ()