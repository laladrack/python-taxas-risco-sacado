import datetime
import os
import csv
import win32com.client
import pandas as pd
import openpyxl

yesterday = datetime.datetime.now() - datetime.timedelta(1)
today = datetime.datetime.now()

def main():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
    path = os.path.join(os.path.expanduser("~"), "Documents")
    path = os.path.join(path, f"{today.strftime("%d-%m-%Y")}")
    if not os.path.exists(path):
        os.mkdir(path)
    for file in os.listdir(path):
        if file.startswith("~"):
            os.remove(file)
    excel = pd.ExcelWriter(path + f"\\taxas {today.strftime("%d-%m-%Y")}.xlsx")

    #excel = path + f"\\taxas {today.strftime("%d-%m-%Y")}.xlsx"
    csv_itau(outlook.GetDefaultFolder(6), path, excel)
    csv_santander(outlook.GetDefaultFolder(6), path, excel)
    excel.close()

#ITAU
def csv_itau(inbox, path, excel):
    #filtra inbox
    messages = inbox.Items
    emails = messages.Restrict("[ReceivedTime] >= '" + today.strftime('%d/%m/%Y') +"'")
    emails.Sort("[ReceivedTime]", True)

    for email in emails:
        if "IBBA Risco Sacado - COTAÇÃO INDICATIVA" in email.Subject:
            body = email.Body

        #acha a tabela de taxas e filtra ate o final dela
            start_index = body.find("Prazo (dias)")
            end_index = body.find("Atenciosamente")
            tabela_email = body[start_index:end_index].strip()
            tabela_email.replace(",", ".")
            tabela = tabela_email.split('\r\n')
            tabela.remove("")
            tabela.remove("Prazo (dias)")
            tabela.remove("Taxa (a.m. linear)")

        #transforma tabela em dicionário -> csv
            dict_tabela = dict({tabela[1]:tabela[3], tabela[5]:tabela[7], tabela[9]:tabela[11], tabela[13]:tabela[15], tabela[17]:tabela[19]})

            with open(path + '\\taxasItau ' + datetime.datetime.now().strftime('%d-%m') + '.csv', 'w', newline='') as myfile:
                wr = csv.DictWriter(myfile, dict_tabela.keys(), delimiter=',')
                wr.writeheader()
                wr.writerow(dict_tabela)

            filein = pd.read_csv(path + '\\taxasItau ' + datetime.datetime.now().strftime('%d-%m') + '.csv')
            filein.to_excel(excel, sheet_name='ITAU + PISO', index=None)

#SANTANDER
def csv_santander(inbox, path, excel):
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    emails = messages.Restrict("[ReceivedTime] >= '" + today.strftime('%d/%m/%Y') +"'")

    for email in emails:
        if "Taxas Faurecia - " in email.Subject:
            attachments = email.Attachments
            for attachment in attachments:
                if attachment.DisplayName.endswith(".xlsx"):
                    excel_banco = path + "\\" + f"taxasSantander {datetime.datetime.now().strftime('%d-%m')}.xlsx"
                    attachment.SaveAsFile(excel_banco)
    
    df = pd.read_excel(excel_banco, engine='openpyxl')
    df = df.drop(columns='DATA')
    df = df[df['Dias Corridos'].isin([30, 45, 60, 75, 90])]
    df.set_index('Dias Corridos',inplace=True)
    df = df.transpose()
    df.to_excel(excel, sheet_name='SANTANDER + PISO', index=None)

if __name__ == "__main__":
    main()