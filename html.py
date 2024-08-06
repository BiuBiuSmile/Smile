
from bs4 import BeautifulSoup

    
    



html_doc = """
<html><head><title>Hello World</title></head>
<body><h2>Test Header</h2>
<p>This is a test.</p>
<a id="link1" href="/my_link1">Link 1</a>
<a id="link2" href="/my_link2">Link 2</a>
<p>Hello, <b class="boldtext">Bold Text</b></p>
</body></html>
"""

soup = BeautifulSoup(html_doc,'html.parser')
print(soup.prettify())

title=soup.title #標題
print(title)
print(title.string)
print(title.text)
h2=soup.find('h2')
h2s=soup.find_all('h2')
print(h2)
print(h2s[0].text)


link = soup.find_all('a')
for item in link :
    url=item.get('href')
    name=item.text
    
    print(url)
    print(name)
    
link2=soup.find(id='link2')
print(link2.get('href'))
print(link2.text)











