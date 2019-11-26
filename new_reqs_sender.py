import pandas as pd 
import numpy as np 

from email.message import EmailMessage
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

import os
import os.path
from os import path
import time

###########~~~~~~~~~~  BEGIN FUNCTIONS  ~~~~~~~~~~###########
'''
find_between( s, first, last )				       #function to find string between two substrings using str.index()
find_between_r( s, first, last )			       #function to find string between two substrings using str.rindex()
add_notes_and_assign(formatted_reqObj_list = [])   #function to prompt user for notes and recruiter name, prepends info to appropriate requisition
file_name_consolidator(source)                     #function that accepts a directory path and returns a list of all .npy files down that path
req_assigner(formatted_reqObj_list = [])           #function that accepts a list of strings that represent requisitions, and sends them all to the appropriate recuiting team
'''
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

def add_notes_and_assign(formatted_reqObj_list = []):

	notes = ""

	#prompt user for adding notes and append to string object 
	for j in range(0, len(formatted_reqObj_list)):
		notes_response = input("\n~~~~ Requisition " + str(j + 1) + " of " + str(len(formatted_reqObj_list)) + " ~~~~\n" + formatted_reqObj_list[j] + '\n' + "Do you want to add notes to the requisition above? (y/n) ")
		if(notes_response == 'y'):
			notes = input("Type in your notes and press enter: ")
		recruiter = input("Who do you want to assign this requisition to? ")

		if(not notes == ""):
			formatted_reqObj_list[j] = "\nASSIGN: " + recruiter + "\n\nNOTES: " + notes + formatted_reqObj_list[j] + "\n\n"
		else:
			formatted_reqObj_list[j] = "\nASSIGN: " + recruiter + "\n\n" + formatted_reqObj_list[j]

	return formatted_reqObj_list

def file_name_consolidator(source):
	unsent_reqs = []
	for root, dirs, files in os.walk(source):
		for file_name in files:
			if file_name in files:
				if file_name.endswith(('.npy')):
					unsent_reqs.append(file_name)
	return unsent_reqs

def req_assigner(formatted_reqObj_list = []):

	'''
	Shirine Tavakol <stavakol@nelsontalentdelivery.com>
	Lara Evans <levans@nelsontalentdelivery.com>
	Isaac Luiz <iluiz@nelsonstaffing.com>
	Tarik Zeid <tzeid@nelsontalentdelivery.com>
	Jai Joshi <jai.joshi@imspeople.com>
	Sharon Harnett <sharnett@nelsontalentdelivery.com>
	Frank Porter <fporter@nelsontalentdelivery.com>
	Khadijah Ahmad <kahmad@nelsontalentdelivery.com>
	'''

	if(not (len(formatted_reqObj_list) == 0)):
		for z in range(0, len(formatted_reqObj_list)):
			req_id_num = find_between(str(formatted_reqObj_list[z]), "Details For Requisition: Google - ", "\nPay:")

			#setup email server for requisition assignment 
			to = ['tzeid@nelsontalentdelivery.com']
			to_ahmedebad = ['stavakol@nelsontalentdelivery.com','kahmad@nelsontalentdelivery.com','levans@nelsontalentdelivery.com','iluiz@nelsonstaffing.com','tzeid@nelsontalentdelivery.com','jai.joshi@imspeople.com','sharnett@nelsontalentdelivery.com','fporter@nelsontalentdelivery.com']
			gmail_user = 'reqassigner@gmail.com'
			gmail_pwd = 'nelsonpassword'

			msg = MIMEMultipart()
			msg['Subject'] = "Google - " + req_id_num + " assignment"
			msg['From'] = gmail_user
			msg['To'] = ','.join(to)

			text = MIMEText(formatted_reqObj_list[z])
			msg.attach(text)

			if(path.exists("C:\\Users\\tzeid\\Desktop\\Nelson\\Skill_Matrices\\" + req_id_num + "_smarttrack_screenshot.png")):
				skill_matrix_png = open("C:\\Users\\tzeid\\Desktop\\Nelson\\Skill_Matrices\\" + req_id_num + "_smarttrack_screenshot.png", 'rb').read()
				image = MIMEImage(skill_matrix_png, name=os.path.basename("C:\\Users\\tzeid\\Desktop\\Nelson\\Skill_Matrices\\" + req_id_num + "_smarttrack_screenshot.png"))
				msg.attach(image)

			s = smtplib.SMTP("smtp.gmail.com",587)
			s.ehlo()
			s.starttls()
			s.ehlo()
			s.login(gmail_user, gmail_pwd)
			s.sendmail(gmail_user, to, msg.as_string())
			s.quit()

		return print("Emails have been sent.")
	return print('There are no jobs to send out, so no emails will be sent.')

###########~~~~~~~~~~  END FUNCTIONS  ~~~~~~~~~~###########

#loads .npy file with list of strings that represent requisitions 
#the list is fed to add_notes_and_assign() which iterates over the list and prepends the user's input for notes and the recrutier name
#the fixed list with the notes and recruiter name is then fed to req_assigner() which faciliates sending emails to the appropriate recruiters 

unsent_reqs = file_name_consolidator("C:\\Users\\tzeid\\Desktop\\Nelson")

for i in range(0, len(unsent_reqs)):
	req_assigner(add_notes_and_assign(np.load(unsent_reqs[i])))
	os.remove(unsent_reqs[i])