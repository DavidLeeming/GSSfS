import re
from requests_html import HTMLSession
import requests
from bs4 import BeautifulSoup
import pandas as pd
Email_list = []
EMAIL_REGEX = r"""(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"""

URL = "www.broughtoninfant.bucks.sch.uk"
URL_check = URL[:8]
if URL_check == 'http://':
    URL = URL
else:
    URL = 'http://' + URL
    print(URL)

page = requests.get(URL)
soup = BeautifulSoup(page.content, "html.parser")
try:
    for link in soup.find_all('a', attrs={'href': re.compile("^http://")}):
        html = link.get('href')
        print(html)
        if '/contact' in html:
            print(link.get('href'))
            new_url = link.get('href')  
except:
    try:
        session = HTMLSession()
        r = session.get(URL)
        r.html.render()
        for re_match in re.findall('/contact(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', r.html.raw_html.decode()):
            print(re_match)
            new_url = URL + str(re_match)
            print(new_url)
    except:
        pass


session = HTMLSession()
r = session.get(new_url)
r.html.render()
for re_match in re.finditer(EMAIL_REGEX, r.html.raw_html.decode()):
    print(re_match.group())
    match = re_match.group(0)
    if str(match[:1]).isalnum() == True:
        if match not in Email_list: 
            Email_list.append(str(match))

print(Email_list)
