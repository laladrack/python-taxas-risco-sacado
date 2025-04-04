import datetime
import csv
import pandas as pd

yesterday = datetime.datetime.now() - datetime.timedelta(1)
today = datetime.datetime.now()

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
            end_index = body.find("Atacado | Mesa Fornecedores")
            tabela_email = body[start_index:end_index].strip()
            tabela_email.replace(",", ".")
            tabela = tabela_email.split('\r\n')
            tabela.remove("")
            tabela.remove("Prazo (dias)")
            tabela.remove("Taxa (a.m. linear)")

        #transforma tabela em dicionário -> csv
            dict_tabela = dict({tabela[1]:tabela[3], tabela[5]:tabela[7], tabela[9]:tabela[11], tabela[13]:tabela[15], tabela[17]:tabela[19]})

            with open(path + '\\taxasItau ' + datetime.datetime.now().strftime('%d-%m-%Y') + '.csv', 'w', newline='') as myfile:
                wr = csv.DictWriter(myfile, dict_tabela.keys(), delimiter=',')
                wr.writeheader()
                wr.writerow(dict_tabela)

            filein = pd.read_csv(path + '\\taxasItau ' + datetime.datetime.now().strftime('%d-%m-%Y') + '.csv')
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
                    excel_banco = path + "\\" + f"taxasSantander {datetime.datetime.now().strftime('%d-%m-%Y')}.xlsx"
                    attachment.SaveAsFile(excel_banco)
    
    df = pd.read_excel(excel_banco, engine='openpyxl')
    
    days_needed = (30, 45, 60, 75, 90)
    
    for day in days_needed:
        found = df[df['Dias Corridos'].astype('str').str.match(f'^{day}$')]
        if len(found) == 0:
            closest_found = False
            while not closest_found:
                for i in range(1, 10):  # limit i to a reasonable number
                    found1 = df[df['Dias Corridos'].astype('str').str.match(f'^{str(day - i)}$')]
                    if len(found1) != 0:
                        df.replace(to_replace=(day - i), value=(day), inplace=True)
                        closest_found = True
                        break

    df = df.drop(columns='DATA')
    df['CUSTO MÊS'] = df['CUSTO MÊS'].map('{:,.4f}'.format)
    df = df[df['Dias Corridos'].isin(days_needed)]
    df.set_index('Dias Corridos', inplace=True)
    df = df.transpose()
    df.to_excel(excel, sheet_name='SANTANDER + PISO', index=None)