# Importing Libraries
from bs4 import BeautifulSoup
from urllib.request import urlopen
import requests
import pandas as pd 
import time
import selenium
from selenium import webdriver

# Loading the existing excel file
headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'}
inputdata = pd.read_excel('output.xlsx', usecols='C')
print(inputdata)

# Loading the HTML content
temp = inputdata
df1 = pd.DataFrame(temp)
links = df1['Links'].tolist()
#print(links)
#print(temp)
#print(type(temp))

# Scraping Data
title, abt_prod, compo, htuse, opinfo = [], [], [], [], []
for link in links:
    
    res = requests.get(link, headers=headers, verify=False)
    print(res.status_code)
    time.sleep(1)
    soup = BeautifulSoup(res.content, "html.parser")
    #print(soup.prettify())
  
    try:
        title.append(soup.title.text)
        print(soup.title.text)
    except:
        title.append("No Data Found")
    try:
        abt_prod.append(soup.find('div', {'id':'about_0'}).get_text())
    except:
        abt_prod.append("No Data Found")
    try:
        compo.append(soup.find('div', {'id':'about_1'}).get_text())
    except:
        compo.append("No Data Found")
    try:
        htuse.append(soup.find('div', {'id':'about_2'}).get_text())
    except:
        htuse.append("No Data Found")
    try:
        opinfo.append(soup.find('div', {'id':'about_3'}).get_text())
    except:
        opinfo.append("No Data Found")
    time.sleep(3)

# Converting Data to DataFrame
d = {'Title':title, 'About Product':abt_prod, 'Composition/How To Use':compo, 'Benefits/How to Use':htuse, 'Other Product Info':opinfo}
df = pd.DataFrame(d)
df

# Saved to excel sheet
df.to_excel("BSDATA.xlsx")