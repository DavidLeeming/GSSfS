import cmd
import re
import sys
import requests
from requests_html import HTMLSession
from bs4 import BeautifulSoup
import pandas as pd

cwd = (str(sys.argv[0][:-13]))
print(cwd)
Email_Dict = {}
EMAIL_REGEX = r"""(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"""
Website_list = []

def process(line):
    Email_list = []
    URL_list = []
    URL = line
    URL_check = URL[:7]
    URL_check_2 = URL[-1]
    if URL_check == 'http://':
        URL = URL
    elif URL_check == 'https:/':
        URL = URL
    else:
        URL = 'http://' + URL
    if URL_check_2 == '/':
        URL = URL[:-1]
    print(URL)
    page = requests.get(URL)
    try:
        session = HTMLSession()
        r = session.get(URL)
        r.html.render()
        for re_match in re.findall(EMAIL_REGEX, r.html.raw_html.decode()):      
            print(re_match.group())
            match = re_match.group(0)
            if str(match[:1]).isalnum() == True:
                if match not in Email_list:
                    Email_list.append(str(match))
    except:
        pass
    try:
        session = HTMLSession()
        r = session.get(URL)
        r.html.render()
        for re_match in re.findall('/[Cc]ontac(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', r.html.raw_html.decode()):
            new_url = URL + str(re_match)
            print(new_url)
            URL_list.append(new_url)
        new_page = requests.get(new_url)
    except:
        pass
    try:
        print(URL_list)
        for URLs in URL_list:
            session = HTMLSession()
            r = session.get(URLs)
            r.html.render()
            for re_match in re.finditer(EMAIL_REGEX, r.html.raw_html.decode()):
                print(re_match.group())
                match = re_match.group(0)
                if str(match[:1]).isalnum() == True:
                    if str(match[-4]) != '.png':
                        if match not in Email_list:
                            Email_list.append(str(match))
        Email_Dict[line] = [line]
        Email_Dict[line].append(Email_list)
        print(Email_Dict)
    except:
        pass


line = 'www.brackenbury.lbhf.sch.uk'
process(line)


file1 = open(str(cwd) + '/urls.txt', 'r')
lines = file1.readlines()

#for line in lines:
    #print(line)
    #process(line)


print(Email_Dict)
df = pd.DataFrame.from_dict(Email_Dict, orient='index', columns=['Website', 'Email'])
filepath = str(cwd) + '/emails.csv'
df.to_csv(filepath, index= False)
print(df)
