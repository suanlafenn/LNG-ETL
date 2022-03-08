import pandas as pd
import win32com.client
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import Series
import numpy as np
from MOCEmail import MOCEmail
from body_parsing import MOCEmailNotificationBody
import json
import dataclasses
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
messages = inbox.Items
emails = []
final_list = [ ]
for message in messages:
    if 'Platts Asia LNG MOC FINALS' in message.Subject or 'Platts APAC LNG MOC FINALS' in message.Subject:
        email = MOCEmail.from_outlook_message(message)
        emails.append(email)
        temp_dict = {}
        temp_dict['Sender'] = email.sender
        temp_dict['Subject'] = email.subject
        temp_dict['Date'] = email.sent_on
        body = email.body
        temp_dict['Bids'] = body.bids
        temp_dict['Offers'] = body.offers
        temp_dict['Trades'] = body.trades
        temp_dict['Withdrawal'] = body.withdrawal
        temp_dict['Exclusions'] = body.exclusion
        final_list.append(temp_dict)

with open("messages.json", "w+") as f:
    for j in emails:
        f.write(json.dumps(dataclasses.asdict(j)))
        f.write("\n")

df = pd.DataFrame(final_list)
patternDel = "(RE:)"
df = df[~df['Subject'].str.contains(patternDel)]
df = df.set_index(['Sender','Subject','Date'])
df=df.stack().reset_index()
df=df[df[0].str.len() != 0]
df=df.explode(0)
df.to_excel("final.xlsx")
print(df.columns)

