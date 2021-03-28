from bs4 import BeautifulSoup
import docx
import convertapi




# snippet to convert docx file to html 


convertapi.api_secret = 'RlUIXZgYfOrdoTVW'
convertapi.convert('html', {
    'File': 'NDA-two-page-template.docx',
    'PlainHtml': 'true'
}, from_format = 'docx').save_files('two_page_tem_from_code.html')



