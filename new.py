import serial
import time
import csv
import math
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime

import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

scopes = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']

credentials = Credentials.from_service_account_file('secret.json', scopes=scopes)

gc = gspread.authorize(credentials)

gauth = GoogleAuth()
drive = GoogleDrive(gauth)

# open a google sheet
gs = gc.open_by_key('1sVpNRP_cFaK2PIAGaoOS7y_FfHOz3JmyDkHzRWKdvKU')
# select a work sheet from its name
worksheet1 = gs.worksheet('Sheet1')

serPort = serial.Serial('COM10', 115200)
serPort.close()
serPort.open()
if (serPort.is_open):
    print ("PORT OPEN: COM10 - 115200 baudrate")
else:
    print ("ERROR")
    exit()

time.sleep(3)
line = ''
prevLine = ''
while True:
  row = []
  while serPort.in_waiting:
    prevLine = line
    line = serPort.readline()
    print(line, prevLine)
    if line == b'P\r\n' and prevLine != b'F\r\n': 
      row = ['pass']
    elif line == b'F\r\n':
      row = ['fail']
    else:
       continue
  
    df = pd.DataFrame({'Size': ['T12'], 'Sender': ['Leadscrew Testing Jig'], 'Pass / Fail': [row[0]], 'Time': datetime.now().strftime("%d/%m/%Y %H:%M:%S")})
    df_values = df.values.tolist()
    gs.values_append('Sheet1', {'valueInputOption': 'RAW'}, {'values': df_values})