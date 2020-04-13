# -*- coding: utf-8 -*-
"""
Created on Mon Apr 13 14:46:33 2020

@author: venug
"""

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from tqdm import tqdm_notebook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import selenium.webdriver.support.expected_conditions as EC
import time
import os




defaultDirectory = r'C:\Users\venug\Downloads\PharmaviewDailyData\HospitalSector'
options = webdriver.ChromeOptions()
# options.add_argument('--headless')
options.add_argument('--incognito')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
# options.add_argument("download.default_directory=C:/Users/venug/Downloads/PharmaviewData/")
prefs = {"profile.default_content_settings.popups": 0,
             "download.default_directory": 
                        r'C:\Users\venug\Downloads\PharmaviewDailyData\HospitalSector\\',#IMPORTANT - ENDING SLASH V IMPORTANT
             "directory_upgrade": True}
options.add_experimental_option('prefs',prefs)
initialCount = len(os.listdir(defaultDirectory))
driver = webdriver.Chrome(r'C:\Users\venug\Downloads\chromedriver.exe',options = options)
driver.get('https://dynamics.pharmaview.dk/ords/f?p=300:101')


driver.find_element_by_id("P101_USERNAME").send_keys('amgsaga')
driver.find_element_by_id("P101_PASSWORD").send_keys('amgsaga')
driver.find_element_by_id('P101_LOGIN').click()
timeout = 500
newCount = initialCount
while(initialCount == newCount):
       locator = '//*[@id="direct_menu_menubar_2i"]'
       WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, locator)))
       driver.find_element_by_id("direct_menu_menubar_2i").click()
       driver.find_element_by_id('show_filters').click()
       time.sleep(2)
       driver.find_element_by_id('P80_SECTORS').send_keys('Hospital Sector')
       timeout = 500
       locator = '//*[@id="15813645532759790_orig"]'
       WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, locator)))
       time.sleep(20)
       # data = driver.page_source
       # dataParsed = BeautifulSoup(data,'lxml')
       # Masterdf = pd.DataFrame()
       driver.find_element_by_id('data_sheet_daily_actions_button').click()
       time.sleep(3)
       driver.find_element_by_id("data_sheet_daily_actions_menu_14i").click()
       time.sleep(10)
       driver.find_element_by_id("data_sheet_daily_download_CSV").click()
       #driver.find_element_by_id("data_sheet_daily_download_CSV").click()
       #driver.find_element(By.XPATH,'//*[@id="data_sheet_daily_download_CSV"]').click()
       time.sleep(30)
       newCount = len(os.listdir(defaultDirectory))
driver.close()
