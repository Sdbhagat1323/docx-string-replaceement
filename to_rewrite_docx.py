from bs4 import BeautifulSoup
import convertapi
import docx

replace_word = "New words"

# import and convert docx file to html with api 

convertapi.api_secret = 'RlUIXZgYfOrdoTVW'
convertapi.convert('html', {
    'File': 'NDA-two-page-template.docx',
    'PlainHtml': 'true'
}, from_format = 'docx').save_files('try_one.html')

#open html file
with open("try_one.html", "wb") as f:
    contents = f.read()
    soup = BeautifulSoup(contents, 'lxml')
    words = soup.find_all(text = re.compile(replace_word))
    print(words)

    
    

    
