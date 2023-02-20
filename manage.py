import pandas as pd
import numpy as np
import math
import requests
from datetime import datetime as dt
import os
from dotenv import load_dotenv

QUANTITY_THRESHOLD = 5
FILE_NAME = 'imspr_materials.xlsx'

xlsx_file = pd.ExcelFile(FILE_NAME)
sheet_names = xlsx_file.sheet_names
slack_webhook_url = os.getenv("SLACK_URL")
slack_message = {'text': f">{dt.now().strftime('%Y년 %m월 %d일 %H:%M')}"}

for idx, sheet in enumerate(sheet_names):
    df = pd.read_excel(FILE_NAME, sheet_name=sheet)
    df = df.drop(0)

    if sheet == 'ProX':
        available_number_column = df.iloc[:,17].astype(float)
    else:
        available_number_column = df.iloc[:,16].astype(float)

    min_value = available_number_column.min()
    available_number = math.floor(min_value)
    is_lacked_idx = np.where(available_number_column < QUANTITY_THRESHOLD)
    available_name = df.iloc[is_lacked_idx[0], 2]

    slack_message['text'] += f"\n*{idx+1}. {available_number}대의 iMSPR-{sheet} 제품이 생산 가능합니다.*"
    slack_message['text'] += f"\n최소 보유 수량 ({QUANTITY_THRESHOLD}개) 이하로 보유하고 있는 부품은 아래와 같습니다."

    for name in available_name.dropna():
        slack_message['text'] += f"\n  - {name}"

    if idx != len(sheet_names)-1:        
        slack_message['text'] += f"\n"
        
response = requests.post(slack_webhook_url, json=slack_message)

if response.status_code != 200:
    raise ValueError('Failed to send Slack message: %s' % response.text)




