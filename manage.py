import pandas as pd
import numpy as np
import math
import requests
from datetime import datetime as dt
import os
from dotenv import load_dotenv
import time

load_dotenv()

def checkNaN(target):
    try:
        return np.isnan(target)
    except:
        return False   

def checkString(target):
    try:
        return type(target) is str
    except:
        return False

slack_webhook_url = os.getenv("SLACK_URL")
file_path = os.getenv("FILE_PATH")
quantity_threshold = int(os.getenv("QUANTITY_THRESHOLD"))
xlsx_file = pd.ExcelFile(file_path)
sheet_names = xlsx_file.sheet_names

while slack_webhook_url == '':
    print('Waiting...')

slack_message = {'text': f">{dt.now().strftime('%Y년 %m월 %d일 %H:%M')}"}
name_column_index = 2
available_number_column_index = 17

for idxSheet, sheet in enumerate(sheet_names):
    df = pd.read_excel(file_path, sheet_name=sheet)
    df = df.drop(0)

    receipt_date_column_index = available_number_column_index + 3
    available_number_column_value = df.iloc[:,available_number_column_index].values

    for idxNumber, number in enumerate(available_number_column_value):
        if checkString(number):
            available_number_column_value[idxNumber] = np.nan

    min_value = available_number_column_value[~np.isnan(available_number_column_value.astype(float))].min()    
    is_lacked_row_index = np.where(available_number_column_value < quantity_threshold)
    available_name = df.iloc[is_lacked_row_index[0], name_column_index].values
    
    try:
        receipt_date = df.iloc[is_lacked_row_index[0], receipt_date_column_index].values
    except:
        receipt_date = np.empty(len(available_name))
        receipt_date[:] = np.nan

    available_number = df.iloc[is_lacked_row_index[0], available_number_column_index].values

    slack_message['text'] += f"\n*{idxSheet+1}. {math.floor(min_value)}대의 iMSPR-{sheet} 제품이 생산 가능합니다.*"
    slack_message['text'] += f"\n최소 보유 수량 ({quantity_threshold}대) 이하로 보유하고 있는 부품은 아래와 같습니다."

    for idxAvailableName, name in enumerate(available_name):
        if checkNaN(name):
            continue
        
        number = available_number[idxAvailableName]

        if checkNaN(number):
            number = 0
        else:
            number = int(np.floor(number))

        date = receipt_date[idxAvailableName]
        if checkNaN(date):
            date = '-'
        
        slack_message['text'] += f"\n  - {name} | {number}대 | 입고예정 : {date}"

    if idxSheet != len(sheet_names)-1:        
        slack_message['text'] += f"\n"
        
response = requests.post(slack_webhook_url, json=slack_message)

if response.status_code != 200:
    raise ValueError('Failed to send Slack message: %s' % response.text)


