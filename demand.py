import pandas as pd
from tabulate import tabulate
import tkinter as tk
from tkinter import filedialog
import os
import sys
import datetime

cwd = os.getcwd()

window = tk.Tk()
window.withdraw()

filetypes = (
    ("Excel files", "*.xlsx"),
    ("All files", "*.*")    
)
file_path = filedialog.askopenfilename(title="Select a material xlsx file", initialdir=cwd, filetypes=filetypes)
window.destroy()

if not file_path:
    print("no file selected")
    sys.exit()

df = pd.read_excel(file_path, sheet_name='재료')
demand_prox = int(input("iMSPR-ProX 목표 대수를 입력하세요: "))
demand_pro = int(input("iMSPR-Pro 목표 대수를 입력하세요: "))
demand_mini = int(input("iMSPR-mini 목표 대수를 입력하세요: "))
demand_tcu = int(input("TCU 목표 대수를 입력하세요: "))

parts_stock = pd.Series(data=df['보유'], dtype=int)
parts_price = pd.Series(data=df['Unit price'], dtype=int)
parts_category = pd.Series(data=df['Category'], dtype=str)
parts_name = pd.Series(data=df['Name'], dtype=str)
parts_product_num = pd.Series(data=df['PN'], dtype=str)

parts_brand = pd.Series(data=df['Brand'], dtype=str)
parts_pk_price = pd.Series(data=df['Pk Price'], dtype=str)
parts_pk_qty = pd.Series(data=df['Pk Q\'ty'], dtype=str)
parts_distributor = pd.Series(data=df['취급처'], dtype=str)

parts_demand_prox = pd.Series(data=df['ProX'], dtype=int) * demand_prox
parts_demand_pro = pd.Series(data=df['Pro'], dtype=int) * demand_pro
parts_demand_mini = pd.Series(data=df['mini'], dtype=int) * demand_mini
parts_demand_tcu = pd.Series(data=df['TCU'], dtype=int) * demand_tcu

shortage = parts_stock - parts_demand_prox - parts_demand_pro - parts_demand_mini
shortage_num = shortage[shortage < 0] * -1
shortage_idx = shortage[shortage < 0].index
shortage_unit_price = parts_price[shortage < 0]
shortage_category = parts_category[shortage_idx]
shortage_name = parts_name[shortage_idx]
shortage_product_num = parts_product_num[shortage_idx]
shortage_price = shortage_unit_price.multiply(shortage_num)
shortage_price_sum = shortage_price.sum()
shortage_price = shortage_price.apply(lambda x: "{:,.0f}원".format(x))

shortage_brand = parts_brand[shortage_idx]
shortage_pk_price = parts_pk_price[shortage_idx]
shortage_pk_qty = parts_pk_qty[shortage_idx]
shortage_distributor = parts_distributor[shortage_idx]
# shortage_unit_price = parts_price[shortage_idx]

df_result = pd.concat([shortage_category,shortage_name,shortage_product_num, shortage_num, shortage_price], axis=1)
# df_result_save = pd.concat([shortage_category,shortage_name,shortage_product_num, shortage_brand, shortage_num, shortage_pk_qty, shortage_pk_price, shortage_distributor], axis=1)
# df_result.rename(columns = {0 : '개수'}, inplace = True)
df_result.rename(columns = {1 : '가격'}, inplace = True)

dict_target = {'iMSPR-ProX' : demand_prox, 
               'iMSPR-Pro' : demand_pro,
               'iMSPR-mini' : demand_mini,
               'TCU module' : demand_tcu}
df_target = pd.DataFrame(dict_target, index=['목표 대수'])

target_table = tabulate(df_target, headers='keys', tablefmt="github")
result_table = tabulate(df_result, headers='keys',tablefmt="github")
print(' ')
print(target_table)
print(' ')
print('- 부족 재료 리스트')
print(result_table)
print('* 필요 금액 : ', shortage_price_sum, '원')

# Save excel file
# Category, Name, Description, PN, Brand, 필요 수량, Pk Price, Pk Q'ty, Unit price, 취급처
df_result_save = pd.concat([shortage_category,shortage_name,shortage_product_num, shortage_brand, shortage_num, shortage_pk_qty, shortage_pk_price, shortage_distributor], axis=1)
df_result_save.rename(columns = {0 : '필요개수'}, inplace = True)

now = datetime.datetime.now()
filename = "{}_demands.xlsx".format(now.strftime("%Y-%m-%d_%H-%M-%S"))
df_result_save.to_excel(filename, index=False)

print("file saved {}".format(filename))

input("Press enter to exit.")