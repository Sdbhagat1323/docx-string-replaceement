import win32com.client

doc = win32com.client.GetObject("docx_InputFile.docx")
doc.SaveAs(FileName="Html_FileName.html", FileFormat=8)
doc.Close()
