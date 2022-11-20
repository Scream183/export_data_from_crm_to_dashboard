#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from time import sleep
import requests as r
from PIL import Image
import win32com.client as win32
from python_anticaptcha import AnticaptchaClient, ImageToTextTask
import pandas as pd
import datetime
import pygsheets
import gspread_dataframe as gd
import os
import schedule
import numpy as np


# In[ ]:


def export_amo(way, login1, password1, url1):
    day = datetime.datetime.now().strftime("%Y-%m-%d")
    file_name = f'\\amocrm_export_leads_{day}.csv'
    way_dowload = f'C:\\Users\\79651\\Downloads{file_name}'
    for i in os.listdir(r'C:\Users\79651\Downloads'):
        if f'amocrm_export_leads_{day}.csv' == i:
            os.remove(way_dowload)
            
    browser = Chrome(way + '\Desktop\chromedriver.exe')
    url = url1
    browser.get(url)
    
    sleep(2)
    
    login = browser.find_element(By.ID,'session_end_login')
    login.send_keys(login1)
    password = browser.find_element(By.NAME,'password')
    password.send_keys(password1)
    
    sleep(1)
    
    browser.find_element(By.ID,'auth_submit').click()
    
    sleep(1)
    
    browser.find_element(By.ID,'search_input').click()
    
    sleep(2)
    
    browser.find_element(By.XPATH,('//*[@id="filter_fields"]/div[3]/div/div/div[2]/span[1]/span/div')).click()
    
    sleep(2)
    
    browser.find_element(By.XPATH,("//div[@title = 'Выбрать всё']")).click()
    
    sleep(2)
    
    browser.find_element(By.XPATH,("//div[@title = 'Выбрать всё']")).click()
    
    sleep(2)
    
    browser.find_elements(By.CLASS_NAME,'button-input-inner ')[2].click()
    
    sleep(5)
    
    browser.find_element(By.XPATH,("//button[@title = 'Еще']")).click()
    
    sleep(2)
    
    browser.find_element(By.XPATH,('//*[@id="list__body-right"]/div[1]/div[3]/div/div/ul/li[3]/div')).click()
    
    sleep(2)
    
    browser.find_element(By.XPATH,("//label[@for = 'export_csv']")).click()
    
    sleep(2)                    
    
    browser.find_element(By.XPATH,('/html/body/div[12]/div[1]/div/div/div[2]/label[2]/div[2]/div[3]/button')).click()
    
    sleep(2)
    
    browser.find_elements(By.XPATH,("//span[@title = ' с активным фильтром']"))[1].click()
    
    sleep(2)
    
    browser.find_elements(By.CLASS_NAME,'button-input_blue')[2].click()
    
    sleep(20)
    
    browser.find_element(By.XPATH,('/html/body/div[13]/div[1]/div/div/div[2]/a/button/span/span')).click()
    
    sleep(10)


# In[ ]:


def export_to_googlesheets(way_dowload):
    df=pd.read_csv(way_dowload)
    df['ЛС в обслуживании'] = df['ЛС в обслуживании'].fillna(0).astype(int)
    df = df.fillna(' ')
    df['Дата закрытия'] = df['Дата закрытия'].replace('не закрыта', '')
    df['Причина отказа'] = df['Этап сделки'].str.split('(',expand=True)[1].str.strip(')')
    df['Этап сделки'] = df['Этап сделки'].str.split('(',expand=True)[0]
    client = pygsheets.authorize(r'C:\Users\79651\Downloads\client_secret_481353108653-sael12rlravq7omb9inlspgn07sailet.apps.googleusercontent.com.json')
    print(client)
    sh = client.open('777')
    wks = sh.sheet1
    wks.clear()
    wks.set_dataframe(df, (1,1), encoding='utf-8', fit=True)


# In[ ]:


day = datetime.datetime.now()
day = day.strftime("%Y-%m-%d")
way = r'C:\Users\79651'
url1 = 'https://domyland1.amocrm.ru/leads/pipeline/3639910/'
login1 = 'a.yusupov@domyland.ru'
password1 = 'Amsterdamm0725'
file_name = f'\\amocrm_export_leads_{day}.csv'
export_amo(way, login1, password1, url1)


# In[ ]:


day = datetime.datetime.now()
day = day.strftime("%Y-%m-%d")
file_name = f'\\amocrm_export_leads_{day}.csv'
way_dowload = f'C:\\Users\\79651\\Downloads{file_name}'
export_to_googlesheets(way_dowload)


# In[ ]:


schedule.every(1).hour.do(export_amo,way = r'C:\Users\79651', login1 = 'a.yusupov@domyland.ru', password1 = 'Amsterdamm0725', url1 = 'https://domyland1.amocrm.ru/leads/pipeline/3639910/')


# In[ ]:


while True:
    schedule.run_pending()


# In[ ]:




