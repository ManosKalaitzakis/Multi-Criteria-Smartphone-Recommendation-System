#!/usr/bin/env python
# -*- coding: utf-8 -*-
from termcolor import colored
import xlsxwriter
workbook =xlsxwriter.Workbook('Output.xlsx')
worksheet = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()
Specs=['RAM','Ισχύς Βασικού Επεξεργαστή','Ανάλυση','Χωρητικότητα','Μέγεθος','Ανάλυση','Συνδεσιμότητα']
def find_value(spec,soup):
    result = soup.find('dt', text=spec)
    #print(result.nextSibling.text)
    if result:
        return result.nextSibling.text
    else:

        return 0

import requests
from bs4 import BeautifulSoup as bs4
import many_stop_words
import unicodedata
import unidecode
import collections
import pandas as pd
import matplotlib.pyplot as plt
import subprocess
import sys
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException
import re
import time
from random import randint
import winsound
urlorigin='https://www.skroutz.gr'
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--incognito')
options.add_argument("--disable-infobars")
options.add_argument("--disable-extensions")
options.add_argument("--start-maximized")
#options.add_argument('--headless')
options.add_argument("window-size=1980,1080")
driver = webdriver.Chrome(chrome_options=options)
driver.implicitly_wait(10)
def strip_accents(s):
   return ''.join(c for c in unicodedata.normalize('NFD', s)
                  if unicodedata.category(c) != 'Mn')
row_counter=1
kritikh_counter=0
thetika_var = ''
arnhtika_var = ''
oudetera_var = ''
while(row_counter<=5):
    jj = randint(1, 2) / 4
    time.sleep(jj)
    print('Κινητό Νο:',row_counter)
    with open('UrlKinhtwn.txt') as f1:
        kinhto = f1.readline().strip()
    f1.close()
    print('\n')
    print(kinhto)
    k=requests.get(kinhto)
    time.sleep(0.2)
    kinhto = kinhto.split(" ")[0]
    driver.get(kinhto)
    time.sleep(0.3)
    winsound.Beep(300, 700)
    time.sleep(5)


    soupk=bs4(k.text,features='html5lib')
    #print(soupk)
    #print (soupk)
    time.sleep(0.2)
    try:
        onoma_kinhtou=(soupk.find("h1", attrs={"class": "page-title"}))
        print(colored("Όνομα", 'red'), ':', onoma_kinhtou.text)
    except:
        winsound.Beep(300, 700)
        input("Press Enter to continue...")
        with open('UrlKinhtwn.txt', 'a') as f1:
            f1.write('\n'+kinhto)
        f1.close()


        pass
    worksheet2.write(row_counter, 0, onoma_kinhtou.text)
    spec_counter=2
    for spec in Specs:
        kk=find_value(spec,soupk)
        print (colored (spec,'red'),':',kk)
        worksheet2.write(row_counter, spec_counter, kk)
        spec_counter += 1
    soupk1=soupk.find("div", {"class":'spec-groups'})

    mpataria=soupk1.find(text=lambda t: "mAh" in t)
    print (print (colored ("Χωρητικότητα Μπαταρίας",'red'),':',mpataria))
    worksheet2.write(row_counter, 8, mpataria)
    kamera=soupk1.find(text=lambda t: "MP" in t)
    print (print (colored ("Βασική Κάμερα",'red'),':',kamera))






    try:

        time.sleep(1.2)

        time.sleep(1.3)
        aa=driver.find_element_by_class_name("sku-reviews-order")
        time.sleep(1)
        aa.click()
        time.sleep(2.2)
        aa=driver.find_element_by_xpath('//a[@href="#helpful"]').click()
        time.sleep(0.5)
        soupk1= driver.page_source
        time.sleep(0.2)
        soupk=bs4(soupk1,features='html5lib')
        time.sleep(1.5)
        onoma_kinhtou=(soupk.find("h1", attrs={"class": "page-title"}))
        print('\n')
        print(colored("Συλλογή Κριτικών για το", 'red'), ':', onoma_kinhtou.text)
        time.sleep(1.2)
        more = driver.find_element_by_class_name("load-more-reviews")
        for load in range(1,18):
            time.sleep(0.3)
            driver.execute_script("arguments[0].click();", more)
            sleeping_time=0.2+ randint(1,3)
            time.sleep(sleeping_time)
            more = driver.find_element_by_class_name("load-more-reviews")
    except:

        winsound.Beep(300, 700)
        input("Press Enter to continue...")
        with open('UrlKinhtwn.txt', 'a') as f1:
            f1.write('\n' + kinhto)
        f1.close()
    finally:

        pass

    soupk1 = driver.page_source
    time.sleep(0.2)
    soupk = bs4(soupk1, features='html5lib')
    x=0
    al = soupk.findAll("li", {"class": ["card review-item","cf media-object faded card review-item","cf media-object faded card review-item expandable","card review-item expandable","card review-item demoted"]})
    l=0
    l1=0
    mhkos=0
    i1=0
    for count in al:
        i1=i1+1
    if i1<90:
        row_counter=row_counter-1
        pass
    for ii in al:
        kategoria=True
        l=l+1
        user=ii.find('div',{'class':'media-text'})
        user_simple=user.text.split('στις')[0]
        print (colored ('Όνομα Χρήστη','red'),':',user_simple)

        n=ii.find('div',{'class':'rating stars'})
        o=n.text             #βαθμολογια
        print (colored ("Βαθμολογία",'red'),':',o,"/ 5")
        k=ii.find('div',{'class':'review-body'})
        text=k.text       #σχολιο_σε_κειμενο

        a_list = text.split()
        text = " ".join(a_list)
        k1 = ii.find('div', {'class': 'review-aggregated-data cf'})
        try:
            k1.findAll('ul', {'class': ['icon pros', 'icon cons', 'icon so-so']})
        except AttributeError:
            print('δεν έχει βάλει κατηγορίες στην κριτική')
            thetika_var = '-'
            arnhtika_var = '-'
            oudetera_var = '-'
            kategoria=False
            pass


        if kategoria==True:
            for thetika in k1.select('ul[class*="icon pros"]'):
                if not thetika.text.strip():
                    thetika_var='--'
                else:
                    for thetika_var1 in thetika:
                        thetika_var=thetika_var+ thetika_var1.text + ','
                    thetika_var=thetika_var[:-1]
            for arnhtika in k1.select('ul[class*="icon cons"]'):
                if not arnhtika.text.strip():
                    arnhtika_var='--'
                else:
                    for arnhtika_var1 in arnhtika:
                        arnhtika_var=arnhtika_var+ arnhtika_var1.text + ','
                    arnhtika_var=arnhtika_var[:-1]
            for oudetera in k1.select('ul[class*="icon so-so"]'):
                if not oudetera.text.strip():
                    oudetera_var='--'
                else:
                    for oudetera_var1 in oudetera:
                        oudetera_var=oudetera_var+ oudetera_var1.text + ','
                    oudetera_var=oudetera_var[:-1]

        if (i1>=30):#len(text) <= 800 kai na exei 25 kritikes
            text=bytes(text, 'utf-8').decode('utf-8', 'ignore')


            if(1):
                print(colored("Κέιμενο Κριτικής", 'red'), ':')
                print(text)
                mhkos = mhkos + len(text)
                l1=l1+1
                '''
                worksheet.write(row_counter, 0, onoma_kinhtou.text)

                if not thetika_var:
                        thetika_var='-'
                if not arnhtika_var:
                        arnhtika_var='-'
                if not oudetera_var:
                        oudetera_var='-'
                if l1==1:
                    worksheet.write(row_counter , l1, text)
                    worksheet.write(row_counter , l1+1, int(o))

                    worksheet.write(row_counter , l1 + 2, thetika_var)
                    worksheet.write(row_counter , l1 + 3, arnhtika_var)
                    worksheet.write(row_counter , l1 + 4, oudetera_var)
                    worksheet.write(row_counter, l1 + 5, user_simple)

                else:
                    worksheet.write(row_counter , (l1+(-1))*6+1, text)
                    worksheet.write(row_counter , (l1 + (-1)) * 6 + 2, int(o))
                    worksheet.write(row_counter , (l1 + (-1)) * 6 + 3, thetika_var)
                    worksheet.write(row_counter , (l1 + (-1)) * 6 + 4, arnhtika_var)
                    worksheet.write(row_counter , (l1 + (-1)) * 6 + 5, oudetera_var)
                    worksheet.write(row_counter, (l1 + (-1)) * 6 + 6, user_simple)


        '''

            kritikh_counter =  l1
        thetika_var = ''
        arnhtika_var = ''
        oudetera_var = ''

    i1=0

    #f.close()
    with open('UrlKinhtwn.txt', 'r') as fin:
        data = fin.read().splitlines(True)
    with open('UrlKinhtwn.txt', 'w') as fout:
        fout.writelines(data[1:])
    row_counter=row_counter+1
workbook.close()