import pandas as pd 
import csv
import numpy as np 
import math

from oauth2client.service_account import ServiceAccountCredentials

import os
import os.path
from os import path
import time

from io import StringIO

import datetime
from datetime import date

PATH = "C:/Users/ltran/Downloads"


st_df = pd.read_csv(PATH + "/xW_List.csv", keep_default_na=False)
er_df = pd.read_csv(PATH + "/Expense Output.csv")

print("-------------------------new-run---------------------------------")

er_df.columns = ['Name', 'num', 'date', 'paid', 'billed', 'reason', 'some label']


st_df['first name'] = st_df['Name'].str.split(", ").str[1]
st_df['last name'] = st_df['Name'].str.split(", ").str[0]

st_df['full name'] = st_df['first name'] + st_df['last name']


er_df['first name'] = er_df['Name'].str.split().str[1]
er_df['last name'] = er_df['Name'].str.split().str[0]

er_df['full name'] = er_df['first name'] + "  " + er_df['last name']

#st_df.set_index('full name', inplace=True)
#st_df.update(er_df.set_index('xW#'))

blah = st_df.merge(er_df, on="full name")

blah = blah[['Name_y', 'num', 'date', 'paid', 'billed', 'reason', 'some label', 'xW#']]
blah.to_csv(PATH + "/Expense Output.csv", index=False, header=False)
