import pandas as pd
import requests

QUANTITY_THRESHOLD = 5
FILE_NAME = 'imspr_materials.xlsx'

xlsx_file = pd.ExcelFile(FILE_NAME)
sheet_names = xlsx_file.sheet_names
print(sheet_names)

df = pd.read_excel(FILE_NAME, sheet_name=sheet_names[0])
print(df)
products_available = df['제작 가능 대수'].min()
print(products_available)
# lacked_materials = df[df['Quantity'] < QUANTITY_THRESHOLD]

# # Check if there are enough materials to produce at least one product
# if products_available == 0 or not lacked_materials.empty:
#     # Send a message to Slack
#     slack_webhook_url = '<slack_webhook_url>'
#     slack_message = {'text': 'Alert: Not enough materials to produce a product!'}
#     if not lacked_materials.empty:
#         slack_message['text'] += '\n\nLacked materials:\n' + lacked_materials.to_string(index=False)
#     response = requests.post(slack_webhook_url, json=slack_message)
#     if response.status_code != 200:
#         raise ValueError('Failed to send Slack message: %s' % response.text)
