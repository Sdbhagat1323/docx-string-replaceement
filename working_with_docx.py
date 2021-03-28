from bs4 import BeautifulSoup
import docx
import convertapi


 
convertapi.api_secret = 'RlUIXZgYfOrdoTVW'
convertapi.convert('docx', {
    'File': 'result.html'
}, from_format = 'html').save_files('from_html_to_doc_tr_code.docx')