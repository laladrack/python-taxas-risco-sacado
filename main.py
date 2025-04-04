import outlookTaxas
import datetime
import os
import win32com.client
import pandas as pd

def main():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
    today = datetime.datetime.now()
    path = os.path.join(os.path.expanduser("~"), "Documents")
    path = os.path.join(path, f"{today.strftime("%d-%m-%Y")}")
    if not os.path.exists(path):
        os.mkdir(path)
    for file in os.listdir(path):
        if file.startswith("~"):
            os.remove(file)
    excel = pd.ExcelWriter(path + f"\\taxas {today.strftime("%d-%m-%Y")}.xlsx")

    #excel = path + f"\\taxas {today.strftime("%d-%m-%Y")}.xlsx"
    outlookTaxas.csv_itau(outlook.GetDefaultFolder(6), path, excel)
    outlookTaxas.csv_santander(outlook.GetDefaultFolder(6), path, excel)
    excel.close()

if __name__ == "__main__":
    main()