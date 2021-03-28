import zipfile
from bs4 import BeautifulSoup

soup = BeautifulSoup("from_code.html", 'lxml')

print(soup)