#!/usr/bin/env python
# coding: utf-8

# In[2]:


import requests
import os
import sys
path = os.path.join(os.path.expanduser("~"), 'Desktop')
import re
import pandas as pd
import datetime
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders("freight") # "6" refers to the index of a folder - in this case,
messages = inbox.Items
email_subject = []
email_date = []
email_content = []

for message in messages:
    if 'FW: BRS CLEAN AG REPORT' in message.Subject:
        email_subject.append(message.Subject)
        email_date.append(message.senton.date()) 
        string = message.body
        lines = string.split('\r\n')
        lines_stripped = [line.strip() for line in string.split('\r\n') if line.strip() != '']
        email_content.append(lines_stripped)
# BRS CLEAN AG REPORT-- LR2\'s
content = []
for email in (email_content):
    if 'LR2\'s' in email:  
        text = email[email.index('FXD/FLD') : email.index('LR1\'s')+1]
        content.append(text)
        
#take body from nested list
body =[]
for sets in content:
    sets = sets[sets.index('FXD/FLD')+1:sets.index('LR1\'s')]
    body.append(sets) 

#write the nested list
body_list = [val for sublist in body for val in sublist]
index_list = [i for i, item in enumerate(body_list) if re.search(r'\d{2}-\d{2} \w{3}', item)]
nested_list = [body_list[s-1:e-1] for s, e in zip([0]+index_list, index_list)]
nested_list = nested_list[1:]
for row in nested_list:
    while len(row) != 8:
        row.append(None)
Charterer = []
load_date = []
Cargo_list = []
Cargo = []
Type = []
Loading_region = []
Discharge_region= []
Vessel_name = []
Freight_rate = []
status = []
for i in nested_list:
    Charterers = i[0]
    Charterer.append(Charterers)
for i in nested_list:
    load_dates = i[1]
    load_date.append(load_dates)
for i in nested_list:
    Cargos = i[2]
    Cargo_list.append(Cargos)
for i in Cargo_list:
    a = i.split(' ')[0]
    b = i.split(' ')[1]
    Cargo.append(a)
    Type.append(b)
for i in nested_list:
    Loading_regions = i[3]
    Loading_region.append(Loading_regions)
for i in nested_list:
    Discharge_regions = i[4]
    Discharge_region.append(Discharge_regions)
for i in nested_list:
    Vessel_names = i[5]
    Vessel_name.append(Vessel_names)
for i in nested_list:
    Freight_rates = i[6]
    Freight_rate.append(Freight_rates)
for i in nested_list:
    statuss = i[7]
    status.append(statuss)


data1 = pd.DataFrame({'Charterer': Charterer,'load_date': load_date,'Cargo': Cargo,'Type': Type,
                    'Loading_region': Loading_region,'Discharge_region': Discharge_region,'Vessel_name': Vessel_name,
                    'Freight_rate': Freight_rate,'status': status})


# BRS CLEAN AG REPORT-- LR1's
content = []
for email in (email_content):
    if 'LR1\'s' in email:  
        text = email[email.index('LR1\'s')+2 : email.index('MR\'s')+1]
        content.append(text)
#take body from nested list
body =[]
for sets in content:
    sets = sets[sets.index('FXD/FLD')+1:sets.index('MR\'s')]
    body.append(sets)  

    
Charterer = []
load_date = []
Cargo_list = []
Cargo = []
Type = []
Loading_region = []
Discharge_region= []
Vessel_name = []
Freight_rate = []
status = []
for i in body:
    Charterers = i[::8]
    Charterer.append(Charterers)
Charterer  = [val for sublist in Charterer for val in sublist]

for i in body:
    load_dates = i[1::8]
    load_date.append(load_dates)
load_date  = [val for sublist in load_date for val in sublist]

for i in body:
    Cargos = i[2::8]
    Cargo_list.append(Cargos)
Cargo_list  = [val for sublist in Cargo_list for val in sublist]


for i in Cargo_list:
    a = i.split(' ')[0]
    b = i.split(' ')[1]
    Cargo.append(a)
    Type.append(b)
    
for i in body:
    Loading_regions = i[3::8]
    Loading_region.append(Loading_regions)
Loading_region  = [val for sublist in Loading_region for val in sublist]
for i in body:
    Discharge_regions = i[4::8]
    Discharge_region.append(Discharge_regions)
Discharge_region  = [val for sublist in Discharge_region for val in sublist]
for i in body:
    Vessel_names = i[5::8]
    Vessel_name.append(Vessel_names)
Vessel_name  = [val for sublist in Vessel_name for val in sublist]
for i in body:
    Freight_rates = i[6::8]
    Freight_rate.append(Freight_rates)
Freight_rate  = [val for sublist in Freight_rate for val in sublist]
for i in body:
    statuss = i[7::8]
    status.append(statuss)
status  = [val for sublist in status for val in sublist]

data2 = pd.DataFrame({'Charterer': Charterer,'load_date': load_date,'Cargo': Cargo,'Type': Type,
                    'Loading_region': Loading_region,'Discharge_region': Discharge_region,'Vessel_name': Vessel_name,
                    'Freight_rate': Freight_rate,'status': status})
# BRS CLEAN AG REPORT-- MR's
content = []
for email in (email_content):
    if 'MR\'s' in email:  
        text = email[email.index('MR\'s')+2 : email.index('Best regards,')+1]
        content.append(text)
#take body from nested list
body =[]
for sets in content:
    sets = sets[sets.index('FXD/FLD')+1:sets.index('Best regards,')]
    body.append(sets) 

#write the nested list
body_list = [val for sublist in body for val in sublist]
body_list = ['00-00 DNR' if x=='DNR-DNR' else x for x in body_list]
index_list = [i for i, item in enumerate(body_list) if re.search(r'\d{2}-\d{2} \w{3}', item)]
nested_list = [body_list[s-1:e-1] for s, e in zip([0]+index_list, index_list)]
nested_list = nested_list[1:]
for row in nested_list:
    while len(row) != 8:
        row.append(None)
Charterer = []
load_date = []
Cargo_list = []
Cargo = []
Type = []
Loading_region = []
Discharge_region= []
Vessel_name = []
Freight_rate = []
status = []
for i in nested_list:
    Charterers = i[0]
    Charterer.append(Charterers)
for i in nested_list:
    load_dates = i[1]
    load_date.append(load_dates)
for i in nested_list:
    Cargos = i[2]
    Cargo_list.append(Cargos)
for i in Cargo_list:
    a = i.split(' ')[0]
    b = i.split(' ')[1]
    Cargo.append(a)
    Type.append(b)
for i in nested_list:
    Loading_regions = i[3]
    Loading_region.append(Loading_regions)
for i in nested_list:
    Discharge_regions = i[4]
    Discharge_region.append(Discharge_regions)
for i in nested_list:
    Vessel_names = i[5]
    Vessel_name.append(Vessel_names)
for i in nested_list:
    Freight_rates = i[6]
    Freight_rate.append(Freight_rates)
for i in nested_list:
    statuss = i[7]
    status.append(statuss)
    


data3 = pd.DataFrame({'Charterer': Charterer,'load_date': load_date,'Cargo': Cargo,'Type': Type,
                    'Loading_region': Loading_region,'Discharge_region': Discharge_region,'Vessel_name': Vessel_name,
                    'Freight_rate': Freight_rate,'status': status})

#save in different spreadsheet
data = pd.concat([data1, data2,data3], ignore_index = True,axis = 0)
data.drop_duplicates(inplace=True)

data['time'] = data['load_date'].astype(str).str.extract(r'((?<=-).*)', expand=True)
data.to_excel(path+'\BRS_CLEAN'+ re.sub(r'[^0-9]','',datetime.datetime.now().strftime("%d%m%Y")) + '.xlsx',index=False)
os.remove("output.xlsx")
print('done')

