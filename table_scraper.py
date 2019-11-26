import os
import os.path
from os import path
import time
import pandas as pd

from io import StringIO

import datetime
from datetime import date

from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

from bs4 import BeautifulSoup

def table_scraper(driver):
	column_labels = ['Requisition#','Job Title','Work Type Level 2','Work Location','Need Date','End Date','Req. Status','#Need(Total)','#Need(Current)','#Submitted','Number of Submission allowed']

	soup = BeautifulSoup(driver.page_source, "html.parser")
	
	data = []
	table = soup.find('table', attrs={'class':"mGrid table dataTable wrapText dt-responsive dtr-inline KNGrid"})

	rows = table.findAll('tr')
	for row in rows:
		cols = row.findAll('td')
		cols = [ele.text.strip() for ele in cols]
		data.append([ele for ele in cols if ele])

	data = data[2:]
	df = pd.DataFrame(data)
	df.columns = column_labels

	return df

def get_num_pages(driver):

	button_checker = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="HistoryTableViewSuppReqs_paginate"]/span[4]')))
	page_label_text = button_checker.text
	page_label_text_parsed = page_label_text.split(' ')
	num_pages = page_label_text_parsed[1]
	return int(num_pages) + 1

def check_exists_by_xpath(xpath):
	try:
		driver.find_element_by_xpath(xpath)
	except NoSuchElementException: 
		return False
	return True

def csv_generator(page_data):
	page_data = page_data.rename({'Number of Submission allowed': 'Submittals Left'}, axis='columns')
	return page_data

#configure chromedriver
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument("--test-type")
options.add_argument("user-data-dir=C:\\Users\\tzeid\\AppData\\Local\\Google\\Chrome\\User Data")
options.add_argument("profile-directory=Tarik")
driver = webdriver.Chrome(executable_path="C:/Users/tzeid/Desktop/Nelson/chromedriver.exe", options=options)
driver.get('https://smarttrack.gcpnode.com/home')
wait = WebDriverWait(driver, 30)

temp_reqs = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="landingPage"]/div[1]/a')))

temp_reqs.click()

time.sleep(40)

table_data = pd.DataFrame()

i = 0
while(i < get_num_pages(driver)):
	if(not i == 0):
		table_data = pd.concat([table_data,table_scraper(driver)], ignore_index = True)
	time.sleep(20)
	next_page = driver.find_element_by_xpath('//*[@id="HistoryTableViewSuppReqs_next"]')
	next_page.click()
	i = i + 1

table_data.to_csv("C:/Users/tzeid/Downloads/RequirementsForSupplierList.csv")