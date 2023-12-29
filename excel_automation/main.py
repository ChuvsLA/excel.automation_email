import os
from datetime import datetime
import pandas as pd
import win32com.client as win32

route = 'base_routes'
files = os.listdir(route)

#create a chart:
consolidated_chart = pd.DataFrame()

#chart information formating:
for file_name in files:
    sale_chart = pd.read_csv(os.path.join(route, file_name))
    #date formatting from excel to pandas' chart:
    sale_chart['Data de Venda'] = pd.to_datetime('01/01/1900') + pd.to_timedelta(sale_chart['Data de Venda'], unit='d')
    #adds sale_chart to consolidated_chart:
    consolidated_chart = pd.concat([consolidated_chart, sale_chart])

#sort dates in order(from older to new):
consolidated_chart = consolidated_chart.sort_values(by='Data de Venda')
#sort chart index in order:
consolidated_chart = consolidated_chart.reset_index(drop=True)
#save consolidated chart in Excel:
consolidated_chart.to_excel('Vendas.xlsx', index=False)

#send it through email:
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = 'python@gmail.com'
today_date = datetime.today().strftime('%d/%m/%Y')
email.Subject = f'Relatório de Vendas {today_date}'
email.Body = f''''
Prezados, 
Segue em anexo o Relatório de Vendas de {today_date} atualizado.
Qualquer coisa estou à disposição.

Um ótimo dia,
Ana Laura.
'''

#attach files:
attachment_route = os.getcwd()
attachment = os.path.join(attachment_route, 'Vendas.xlsx')
email.Attachments.Add(attachment)

email.Send()