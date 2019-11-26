import os
import time


while True:
	os.system('taskkill /F /IM chrome.exe')
	os.system('taskkill /F /IM chromedriver.exe')
	os.system('python C:\\Users\\tzeid\\Desktop\\Nelson\\table_scraper.py')
	os.system('python C:\\Users\\tzeid\\Desktop\\Nelson\\new_reqs_generator.py')
	time.sleep(1200)