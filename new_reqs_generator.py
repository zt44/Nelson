import pandas as pd 
import csv
import numpy as np 
import math

import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe

from oauth2client.service_account import ServiceAccountCredentials

import os
import os.path
from os import path
import time

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

###########~~~~~~~~~~  BEGIN FUNCTIONS  ~~~~~~~~~~###########

'''
next_available_row(worksheet)                      #get earliest empty row in sheet / row considered empty if cell in column 3 is empty
csv_downloader(driver, wait)                       #navigate smarttrack with chromedriver to click "Save As" -> downloads .csv of job requisitions 
req_finder(req_id, driver, wait)                   #navigates smarttrack with chromedriver to find all info for requisition
req_formatter(dataframe)                           #accepts dataframe of newly released reqs and formats each one
sheet_updater(raw_open_data, raw_smarttrack_data)  #in open orders google doc, updates sheets 'open' and 'don't work' with current smarttrack information
find_between( s, first, last )				       #find string between two substrings using str.index()
find_between_r( s, first, last )			       #find string between two substrings using str.rindex()
truncate(number, digits)					       #truncate a number to a specified number 
calc_months_btwn_dates(date1, date2)			   #to calculate the amount of time between two dates in months
list_diff(list1, list2)                            #to determine difference between two lists 
check_exists_by_xpath(xpath)                       #to determine if an element exists via XPATH
'''

def next_available_row(worksheet):
	str_list = list(filter(None, worksheet.col_values(3)))
	return int(len(str_list)+1)

def csv_downloader(driver, wait):
	try:
		element2 = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="landingPage"]/div[1]/a')))
		element2.click()
	except: 
		print("Page didn't load in time.")
	time.sleep(30)
	try:
		element3 = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="HistoryTableViewSuppReqs_wrapper"]/div[2]/div[1]/a/span')))
		element3.click()
	except:
		print("Page didn't load in time.")
	time.sleep(30)
	print("\n")
	return print("List of requisitions successfully downloaded from Smart Track")

def req_finder(req_id, reqObj_list, driver, wait):
	driver.get('https://smarttrack.gcpnode.com/home')
	time.sleep(10)
	try:
		element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="landingPage"]/div[1]/a')))
	except: 
		print("Page didn't load in time.")

	time.sleep(20)

	try:
		# type text
		tempReqs = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="landingPage"]/div[1]/a')))
		tempReqs.click()
	except: 
		print("Page didn't load in time.")

	time.sleep(15)

	req_num = req_id

	try:
		req_num_search = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ReqNo"]')))
		req_num_search.send_keys(req_num) 
	except: 
		print("Page didn't load in time.")

	time.sleep(15)

	try: 
		click_req = driver.find_element_by_link_text(req_num)
		click_req.click()
	except: 
		print("Page didn't load in time.")

	time.sleep(5)

	jd_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="accordion"]/div[1]/a/h3')))
	jd_button.click()

	actions = ActionChains(driver)

	soup = BeautifulSoup(driver.page_source, "html.parser")
	soup = str(soup.get_text())

	#print(soup)
	#for testing purposes writing soup to a file
	
	with open("smarttrack_soup.txt", "w") as f:
		f.write(soup)

	ID = find_between(soup, "Requisition#", "Work Location")
	JD = find_between(soup, "Position Description", "Skill/Experience/Education")

	markup = float(find_between(soup, "Markup", "%"))
	temp_markup = (markup/100) + 1

	minPay = find_between(soup, "\nMin Bill Rate", "\nMax Bill Rate").split(' ')
	minPay = float(minPay[0])/temp_markup
	minPay = truncate(minPay, 2)
	minPay = math.ceil(minPay)

	maxPay = find_between_r(soup, "\n\nMax Bill Rate", "\n\nEquivalent Rate for Full Time Employee").split(' ')
	maxPay = float(maxPay[0])/temp_markup
	maxPay = truncate(maxPay, 2)
	maxPay = math.floor(maxPay)

	PC = find_between(soup, "Program Coordinator", "Driving Required for Assignment?")

	location = find_between(soup, "Work Location", "Desired Start Date")

	title = find_between_r(soup, "\n\nTitle", "  Position Description")

	start_date = find_between(soup, "Desired Start Date", "End Date")
	end_date = find_between_r(soup, "End Date", "Headcount")

	length = calc_months_btwn_dates(start_date, end_date)

	required_skills = find_between(soup, "   Mandatory", "   Desired")

	desired_skills = find_between_r(soup, "  Desired", "\n                            $(document).ready(function () {")

	#scroll to skill matrix and take a screenshot
	if(not ("No records found" in find_between_r(soup, "\n                Skill Matrix", "\n                Certifications"))):
		skill_matrix = driver.find_element_by_xpath('//*[@id="accordion"]/div[1]/div/div[5]/div[7]/div/label[1]')
		actions.move_to_element_with_offset(skill_matrix, 0, 0).perform()
		screenshot = driver.save_screenshot("C:\\Users\\tzeid\\Desktop\\Nelson\\Skill_Matrices\\" + ID + "_smarttrack_screenshot.png")

	reqObj_list.append(Requisition(ID, JD, minPay, maxPay, location, PC, markup, length, title, required_skills, desired_skills))

	return print("Requisition " + str(req_id) + ": " + str(title) + " info successfully collected from Smart Track")

def req_formatter(reqObj_list = []):
	formatted_reqObj_list = []

	for i in range(0, len(reqObj_list)):
		payload_str = '\n' + "Markup: " + str(reqObj_list[i].markup) + "%" + "\n\nDetails For Requisition: " + "Google - " + str(reqObj_list[i].ID) + "\nPay: " + str(reqObj_list[i].minPay) + "/hr" + " - " + str(reqObj_list[i].maxPay) + "/hr" + '\n' + "Location: " + str(reqObj_list[i].location) + "\n\n" + "Contract Length: " + str(reqObj_list[i].length) + " months" + '\n' + '\n' + "Position Title: " + str(reqObj_list[i].title) + '\n' + '\n' + "Job Description:\n" + str(reqObj_list[i].JD) + "\n\nRequired Skills:\n" + str(reqObj_list[i].required_skills) + '\n' + "\n\nDesired Skills:\n" + str(reqObj_list[i].desired_skills) + "\n\n--------------------------------------------" + "\n\n"
		formatted_reqObj_list.append(payload_str)
		formatted_reqObj_list[i].replace("’","'").replace("●", "-").replace("➔", "->").replace("“", '"').replace("”", '"').replace("–", "-").replace("", ">>")

	'''notes = ""

	#prompt user for adding notes and append to string object 
	for j in range(0, len(formatted_reqObj_list)):
		notes_response = input("\n~~~~ Requisition " + str(j + 1) + " of " + str(len(formatted_reqObj_list)) + " ~~~~\n" + formatted_reqObj_list[j] + '\n' + "Do you want to add notes to the requisition above? (y/n) ")
		if(notes_response == 'y'):
			notes = input("Type in your notes and press enter: ")
		recruiter = input("Who do you want to assign this requisition to? ")

		if(not notes == ""):
			formatted_reqObj_list[j] = "\nASSIGN: " + recruiter + "\n\nNOTES: \n" + notes + formatted_reqObj_list[j]
		else:
			formatted_reqObj_list[j] = "\nASSIGN: " + recruiter + "\n\n" + formatted_reqObj_list[j]'''
	file_name = str(datetime.datetime.now().time()).replace(".","_").replace(":","_") + '_new_requisitions'

	if(not (len(reqObj_list) == 0)):
		return np.save(file_name, formatted_reqObj_list)
	else:
		return print("No .npy file will be saved.")

def truncate(number, digits) -> float:
    stepper = 10.0 ** digits
    return math.trunc(stepper * number) / stepper

def sheet_updater(raw_open_data, raw_smarttrack_data):
	list_disappeared = []
	list_req_num_open = []
	list_req_num_smarttrack = []

    #determine disappeared reqs before updating the spreadsheet values 
	latest_wkst_open = spr.worksheet("Open")
	latest_df_wkst_open = pd.DataFrame(latest_wkst_open.get_all_values())
	latest_df_wkst_open.rename(columns=latest_df_wkst_open.iloc[0], inplace=True)
	latest_df_wkst_open.drop(latest_df_wkst_open.index[0], inplace=True)

	latest_df_wkst_open_only_google = latest_df_wkst_open[latest_df_wkst_open['Requisition#'].str.contains('G-REQ')]

	list_req_num_open = list(latest_df_wkst_open_only_google['Requisition#'])

	list_req_num_smarttrack = list(raw_smarttrack_data['Requisition#'])
	
	list_disappeared = list_diff(list_req_num_open, list_req_num_smarttrack)

	latest_wkst_open.clear()
	set_with_dataframe(latest_wkst_open, latest_df_wkst_open, row=1, include_index=False, include_column_header=True, resize=False, allow_formulas=False)

	for y in latest_df_wkst_open.index:
		for x in range(0, len(list_disappeared)):
			if(latest_df_wkst_open.at[y, "Requisition#"] == list_disappeared[x]):
				latest_df_wkst_open.at[y, "Req. Status"] = "Disappeared"

	open_data = pd.DataFrame(wkst_open.get_all_values())
	open_data.rename(columns=open_data.iloc[0], inplace=True)
	open_data.drop(open_data.index[0], inplace=True)

	raw_open_data = open_data

	raw_open_data_only_closed = raw_open_data[raw_open_data['Req. Status'].str.contains('Closed')]
	raw_open_data = raw_open_data[~raw_open_data['Req. Status'].str.contains('Closed')]

	raw_open_data = raw_open_data.rename({'Number of Submission allowed': 'Submittals Left'}, axis='columns')
	raw_smarttrack_data = raw_smarttrack_data.rename({'Number of Submission allowed': 'Submittals Left'}, axis='columns')
	raw_open_data.set_index('Requisition#', inplace=True)
	raw_open_data.update(raw_smarttrack_data.set_index('Requisition#'))
	raw_open_data = raw_open_data.reset_index()

	raw_open_data = raw_open_data[['Date','Client','Requisition#','Job Titles','Min Pay Rate','Max Pay Rate','Location','Length(months)','Req. Status','Submittals Left','#Submitted','#Need(Total)','Verified','Assigned','Recruiter','PC','Notes','Team','Work Type','Work Type Category']]
	wkst_open.clear()
	set_with_dataframe(wkst_open, raw_open_data, row=1, include_index=False, include_column_header=True, resize=False, allow_formulas=False)

	next_dontwork_row = next_available_row(wkst_open_dontwork)
	wkst_open_after = spr.worksheet("Open")
	wkst_open_after_df = pd.DataFrame(wkst_open_after.get_all_values())
	wkst_open_after_df.rename(columns=wkst_open_after_df.iloc[0], inplace=True)
	wkst_open_after_df.drop(wkst_open_after_df.index[0], inplace=True)
	wkst_open_after_df_without_filled = wkst_open_after_df[~wkst_open_after_df['Req. Status'].str.contains('Filled')]
	wkst_open_after_df_only_filled = wkst_open_after_df[wkst_open_after_df['Req. Status'].str.contains('Filled')]

	frames = [wkst_open_after_df_only_filled, raw_open_data_only_closed]

	wkst_open_after_df_only_filled_and_closed = pd.concat(frames)

	wkst_open.clear()
	set_with_dataframe(wkst_open, wkst_open_after_df_without_filled, row=1, include_index=False, include_column_header=True, resize=False, allow_formulas=False)
	set_with_dataframe(wkst_open_dontwork, wkst_open_after_df_only_filled_and_closed, row=next_dontwork_row, include_index=False, include_column_header=False, resize=False, allow_formulas=False)


	return print("Current requisition information has been updated.")

def find_between(s, first, last):
    try:
        start = s.index( first ) + len( first )
        end = s.index( last, start )
        return s[start:end].strip()
    except ValueError:
        return ""

def find_between_r(s, first, last):
    try:
        start = s.rindex( first ) + len( first )
        end = s.rindex( last, start )
        return s[start:end].strip()
    except ValueError:
        return ""

def calc_months_btwn_dates(date1, date2):
	d1 = pd.to_datetime(date1)
	d2 = pd.to_datetime(date2)
	length = d2 - d1 
	length = round(length/np.timedelta64(1,'M'))
	return length

def list_diff(list1, list2): 

	return (list(set(list1) - set(list2))) 

def check_exists_by_xpath(xpath):
	try:
		driver.find_element_by_xpath(xpath)
	except NoSuchElementException: 
		return False
	return True

###########~~~~~~~~~~  END FUNCTIONS  ~~~~~~~~~~###########

#class definition of a requistion
class Requisition: 
	def __init__(self, ID, JD, minPay, maxPay, location, PC, markup, length, title, required_skills, desired_skills):
		self.ID = ID
		self.JD = JD
		self.minPay = minPay
		self.maxPay = maxPay
		self.location = location
		self.PC = PC
		self.markup = markup
		self.length = length
		self.title = title
		self.required_skills = required_skills
		self.desired_skills = desired_skills

#configure chromedriver
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument("--test-type")
#options.add_argument("--incognito")
options.add_argument("user-data-dir=C:\\Users\\tzeid\\AppData\\Local\\Google\\Chrome\\User Data")
options.add_argument("profile-directory=Tarik")
driver = webdriver.Chrome(executable_path="C:/Users/tzeid/Desktop/Nelson/chromedriver.exe", options=options)
driver.get('https://smarttrack.gcpnode.com/home')
wait = WebDriverWait(driver, 20)

#download smarttrack info
#csv_downloader(driver, wait)

#get current date to paste into spreadsheet rows
now = datetime.datetime.now()
date_today = str(now.month) + str('/') + str(now.day) + str('/') + str(now.year)

#load csv file as pandas object
try: 
	f = pd.read_csv("C:/Users/tzeid/Downloads/RequirementsForSupplierList.csv", keep_default_na=False)
except IOError: 
	f = pd.DataFrame()
	print("File not found.")

#calculate time difference in months and add column to dataframe
df = pd.DataFrame(f)
raw_smarttrack_data = pd.DataFrame(f)
#print(raw_smarttrack_data)

d1 = pd.to_datetime(df["Need Date"])
d2 = pd.to_datetime(df["End Date"])

df['diff_months'] = d2 - d1 

df['diff_months'] = round(df['diff_months']/np.timedelta64(1,'M'))

df["Min Pay Rates"] = " "

#define names of columns that we want to keep in our dataset
keep_col = ['Requisition#', 'Job Title', 'Min Pay Rates', 'Work Type Level 2', 'Work Location', 'diff_months', 'Req. Status', '#Need(Total)', '#Submitted', 'Number of Submission allowed']

#load modified pandas object into new variable
f_correct_columns = f[keep_col]

f_correct_columns.rename(columns={"Number of Submission allowed": "Submittals Left"})

#filter out unwanted requisitions
f_filtered_fiber = f_correct_columns[~f_correct_columns['Job Title'].str.contains('Fiber')]
f_filtered_phoenix = f_filtered_fiber[~f_filtered_fiber['Job Title'].str.contains('Phoenix')]
f_filtered_interviews = f_filtered_phoenix[~f_filtered_phoenix['Req. Status'].str.contains('Interviewing')]
f_filtered_declined = f_filtered_interviews[~f_filtered_interviews['Req. Status'].str.contains('Declined')]
f_filtered_filled = f_filtered_declined[~f_filtered_declined['Req. Status'].str.contains('Filled')]
f_filtered_sourcer = f_filtered_filled[~f_filtered_filled['Job Title'].str.contains('Sourcer')]
f_corrected = f_filtered_sourcer

#define gspread authentications and links
scope = ['https://www.googleapis.com/auth/drive', 'https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/tzeid/Desktop/Nelson/Write to Sheets-e2cd9fc6057c.json', scope)
gc = gspread.authorize(credentials)
spr = gc.open_by_url('https://docs.google.com/spreadsheets/d/1D1ZIIqbGgGKzMpqK4XnyPYc_PtgeJUFnsroIjD2R9g4/edit#gid=1826710698')

#get worksheets to edit 
wkst_open = spr.worksheet("Open")
wkst_open_dontwork = spr.worksheet("Don't Work")
wkst_ims_req_assign = spr.worksheet("IMS Req Assign") 
#wkst_tester = spr.worksheet("tester")

#fix column labels for open, don't work, and ims req assign
open_data = pd.DataFrame(wkst_open.get_all_values())
open_data.rename(columns=open_data.iloc[0], inplace=True)
open_data.drop(open_data.index[0], inplace=True)
raw_open_data = open_data
open_data = open_data[['Requisition#', 'Job Titles']]

dontwork_data = pd.DataFrame(wkst_open_dontwork.get_all_values())
dontwork_data.rename(columns=dontwork_data.iloc[0], inplace=True)
dontwork_data.drop(dontwork_data.index[0], inplace=True)
dontwork_data = dontwork_data[['Requisition#', 'Job Titles']]

assignment_protocol_data = pd.DataFrame(wkst_ims_req_assign.get_all_values())
assignment_protocol_data.rename(columns=assignment_protocol_data.iloc[0], inplace=True)
assignment_protocol_data.drop(assignment_protocol_data.index[0], inplace=True)

print("Comparing to Open Orders spreadsheet...")

#make sure data from .csv is seen as dataframe
df_new = pd.DataFrame(f_corrected)

#merge 'Don't Work' data with data from .csv file
df_merged_with_dontwork = pd.merge(dontwork_data, df_new, how='outer', indicator=True)
df_merged_with_dontwork.drop_duplicates('Requisition#', keep=False, inplace=True)
df_corrected = df_merged_with_dontwork.loc[df_merged_with_dontwork._merge == 'right_only', keep_col]

#merge 'Open' data with data from .csv file and 'Don't Work' data
df_merged_with_open = pd.merge(open_data, df_corrected, how='outer', indicator=True)
df_merged_with_open.drop_duplicates('Requisition#', keep=False, inplace=True)
df_corrected_final = df_merged_with_open.loc[df_merged_with_open._merge == 'right_only', keep_col]
 
#calculate next available row on open sheet
next_row = int(next_available_row(wkst_open))

#declare dataframe for date and google labels
date_and_google = pd.DataFrame()

num_rows_to_create = df_corrected_final.shape[0]

if(num_rows_to_create == 1):
	print("There is", num_rows_to_create, "job that needs sourcing!")
else:
	print("There are", num_rows_to_create, "jobs that need sourcing!")

#create smaller dataframe to fill in other data on sheet ('Google' and 'Date' columns)
for i in range(0, num_rows_to_create):
	date_and_google = date_and_google.append(pd.Series([date_today, "Google"]), ignore_index=True)

#pull work type labels to put somwhere else on spreadsheet
work_type_label = pd.DataFrame(df_corrected_final["Work Type Level 2"])

#define list to store Requisition objects
reqObj_list = []

#add dummy columns to adjust position of data on google sheet (column labels irrelevent) 
df_corrected_final['col1'] = ""
df_corrected_final['col2'] = ""
df_corrected_final['col3'] = ""
df_corrected_final['col4'] = ""

#delete "Work Type Level 2" stuff 
#and parse stuff 
for i in df_corrected_final.index:
	df_corrected_final.at[i, "Work Type Level 2"] = " "
	str_list = df_corrected_final.at[i, "Work Location"].split("-")
	str_list2 = str_list[1].split(")")
	df_corrected_final.at[i, "Work Location"] = str_list2[0]
	req_finder(df_corrected_final['Requisition#'][i], reqObj_list, driver, wait)

for j in df_corrected_final.index:
	for k in range(len(reqObj_list)):
		if(str(df_corrected_final.at[j,"Requisition#"]) == str(reqObj_list[k].ID)):
			df_corrected_final.at[j, "Min Pay Rates"] = reqObj_list[k].minPay
			df_corrected_final.at[j, "Work Type Level 2"] = reqObj_list[k].maxPay
			df_corrected_final.at[j, "col4"] = reqObj_list[k].PC

#if dataframe is non-empty, paste dataframes for newly added requisitions
#if dataframe is empty, print 'Not appending any jobs to spreadsheet.'
if(df_corrected_final.empty):
	print('Not appending any jobs to spreadsheet.')
else:
	print('Appending jobs to spreadsheet...')
	#print(df_corrected_final)
	set_with_dataframe(wkst_open, date_and_google, row=next_row, col=1, include_index=False, include_column_header=False, resize=False, allow_formulas=False)
	set_with_dataframe(wkst_open, df_corrected_final, row=next_row, col=3, include_index=False, include_column_header=False, resize=False, allow_formulas=False)
	set_with_dataframe(wkst_open, work_type_label, row=next_row, col=19, include_index=False, include_column_header=False, resize=False, allow_formulas=False)

sheet_updater(raw_open_data, raw_smarttrack_data)
req_formatter(reqObj_list)

#delete old requisition list .csv file
os.remove("C:/Users/tzeid/Downloads/RequirementsForSupplierList.csv")