import datetime
import os
import pandas as pd

today = datetime.datetime.now()

path = os.path.join(os.path.expanduser("~"), "Documents")
path = os.path.join(path, f"{today.strftime("%d-%m-%Y")}")

excel = path + f"\\taxas {today.strftime("%d-%m-%Y")}.xlsx"

itau = pd.read_excel(excel, sheet_name="ITAU + PISO", engine='openpyxl')
santander = pd.read_excel(excel, sheet_name="SANTANDER + PISO", engine='openpyxl')