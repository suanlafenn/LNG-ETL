#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests
import os
import sys
path = os.path.join(os.path.expanduser("~"), 'Desktop')
import re
import os
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
    if 'FW: CPP WEST MARKET UPDATE' in message.Subject:
        email_subject.append(message.Subject)
        email_date.append(message.senton.date()) 
        string = message.body
        lines = string.split('\r\n')
        lines_stripped = [line.strip() for line in string.split('\r\n') if line.strip() != '']
        email_content.append(lines_stripped)
#CPP WEST MARKET UPDATE - MED HANDIES
content = []
for email in (email_content):
    if 'MED HANDIES' in email:  
        text = email[email.index('MED HANDIES')+2 : email.index('MR')]
        content.append(text)
body =[]
#take body from nested list 
for sets in content:
    set_index = sets.index('OUTSTANDING CARGOES')
    sets = sets[:set_index]
    for item in sets:
        if len(item) > 100:
            body.append(item)



Vessel_name = []
Cargo = []
other_part = []
Type = []
Loading_port = []
Origin_region = []
load_date = []
Freight_rate = []
Charterer = []
Status = []
for i in body:
    Vessel_name.append(i[0:18].rstrip())
    Cargo.append(i[18:21].strip())
    other_part.append(i[21:])
    Origin_region.append(i[45:67].rstrip()) 
    Charterer.append(i[108:124].rstrip()) 
    Status.append(i[124:].rstrip()) 
    
for i in other_part:
    words = re.findall(r'\S+',i)
    Type.append(words[0])
    word_contain_two = i[:24]
    port = word_contain_two.split(None,1)
    Loading_port.append(port[1].rstrip())
    
for i in body:
    words = i[68:108].split(None,1)
    load_date.append(words[0])
    Freight_rate.append(words[1].rstrip())
    
data_1 = pd.DataFrame({'Vessel_name': Vessel_name,'Cargo': Cargo,'Type': Type,'Loading_region': Loading_port,
                    'Discharge_region': Origin_region,'load_date': load_date,'Freight_rate': Freight_rate,'Charterer': Charterer,
                    'Status': Status})


#CPP WEST MARKET UPDATE - MR
content = []
for email in (email_content):
    if 'MR' in email:
        text = email[email.index('MR')+2 : email.index('LR1')]
        content.append(text)
#take body from nested list

body_2 =[]
for sets in content:
    set_index = sets.index('OUTSTANDING CARGOES:')
    sets = sets[:set_index]
    for item in sets:
        if len(item) > 100:
            body_2.append(item)
            
Vessel_name = []
Cargo = []
other_part = []
Type = []
Loading_port = []
Origin_region = []
load_date = []
Freight_rate = []
Charterer = []
Status = []
for i in body_2:
    Vessel_name.append(i[0:22].rstrip())
    Cargo.append(i[22:25].strip())
    other_part.append(i[25:])
    Origin_region.append(i[49:71].rstrip()) 
    Charterer.append(i[108:119].rstrip()) 
    Status.append(i[120:].rstrip()) 
    
for i in other_part:
    words = re.findall(r'\S+',i)
    Type.append(words[0])
    word_contain_two = i[:24]
    port = word_contain_two.split(None,1)
    Loading_port.append(port[1].rstrip())
    
for i in body_2:
    words = i[71:108].split(None,1)
    load_date.append(words[0])
    Freight_rate.append(words[1].rstrip())
    
data_2 = pd.DataFrame({'Vessel_name': Vessel_name,'Cargo': Cargo,'Type': Type,'Loading_region': Loading_port,
                    'Discharge_region': Origin_region,'load_date': load_date,'Freight_rate': Freight_rate,'Charterer': Charterer,
                    'Status': Status})



#CPP WEST MARKET UPDATE - LR1
content = []
for email in (email_content):
    if 'LR1' in email:
        text = email[email.index('LR1')+2 : email.index('LR2')]
        content.append(text)
#take body from nested list
body_3 =[]
for sets in content:
    set_index = sets.index('OUTSTANDING CARGOES')
    sets = sets[:set_index]
    for item in sets:
        if  len(item) > 100:
            body_3.append(item)

Vessel_name = []
Cargo = []
other_part = []
Type = []
Loading_port = []
Origin_region = []
load_date = []
Freight_rate = []
Charterer = []
Status = []
for i in body_3:
    Vessel_name.append(i[0:19].rstrip())
    Cargo.append(i[19:25].strip())
    Type.append(i[25:50].split(None,1)[0])
    Loading_port.append(i[25:55].split(None,1)[1].rsplit(None,1)[0])
    Origin_region.append(i[47:75].rsplit(None,1)[0].lstrip()) 
    load_date.append(i[65:81])
    Freight_rate.append(i[66:98].split(None,1)[1].lstrip())
    Charterer.append(i[98:120])
    Status.append(i[107:])

    
data_3 = pd.DataFrame({'Vessel_name': Vessel_name,'Cargo': Cargo,'Type': Type,'Loading_region': Loading_port,
                    'Discharge_region': Origin_region,'load_date': load_date,'Freight_rate': Freight_rate,'Charterer': Charterer,
                    'Status': Status})


data_3['load_date'] = data_3['load_date'].str.extract(r'(.*(?=WS|LS|RNR))', expand=True)
data_3['load_date'] = data_3['load_date'].str.replace("E", "").str.strip()
data_3['Status'] = data_3['Status'].str.extract(r'((?<= ).*)', expand=True)
data_3['Status'] = data_3['Status'].str.strip()
data_3['Charterer'] = data_3['Charterer'].str.extract(r'(.*(?=SUBS|FXD|FLD))', expand=True)
data_3['Charterer'] = data_3['Charterer'].str.strip()
data_3['Discharge_region'] = data_3['Discharge_region'].str.replace("M  WAF/UKC", "WAF/UKC").str.strip()
data_3['Discharge_region'] = data_3['Discharge_region'].str.replace("M  WAF", "WAF").str.strip()

#CPP WEST MARKET UPDATE - LR2
content = []
for email in (email_content):
    if 'LR2' in email:
        text = email[email.index('LR2')+2 : email.index('DISCLAIMER')]
        content.append(text)
#take body from nested list

body_4 =[]
for sets in content:
    set_index = sets.index('OUTSTANDING CARGOES')
    sets = sets[:set_index]
    for item in sets:
        if  len(item) > 100:
            body_4.append(item)
body_4 = [elements for elements in body_4 if '1-60 DAYS STORAGE DELIVERY LOME' not in elements]

Vessel_name = []
Cargo = []
other_part = []
Type = []
Loading_port = []
Origin_region = []
load_date = []
Freight_rate = []
Charterer = []
Status = []
for i in body_4:
    Vessel_name.append(i[0:19].rstrip())
    Cargo.append(i[19:25].strip())
    Type.append(i[25:50].split(None,1)[0])
    Loading_port.append(i[25:55].split(None,1)[1].rsplit(None,1)[0])
    Origin_region.append(i[47:75].rsplit(None,1)[0].lstrip()) 
    load_date.append(i[65:98].split(None)[0])
    Freight_rate.append(i[65:98].split(None,1)[1].lstrip())
    Charterer.append(i[98:117].rsplit(None,1)[0].strip())
    Status.append(i[107:])

    
data_4 = pd.DataFrame({'Vessel_name': Vessel_name,'Cargo': Cargo,'Type': Type,'Loading_region': Loading_port,
                    'Discharge_region': Origin_region,'load_date': load_date,'Freight_rate': Freight_rate,'Charterer': Charterer,
                    'Status': Status})
data_4['Status'] = data_4['Status'].str.extract(r'((?<= ).*)', expand=True)
data_4['Status'] = data_4['Status'].str.replace("ENERGY", "").str.strip()
data_4['Charterer'] = data_4['Charterer'].str.replace("FXD", "")
data_4['Charterer'] = data_4['Charterer'].str.replace("FLD", "")
data_4['Charterer'] = data_4['Charterer'].str.replace("SUBS –", "")
data_4['Discharge_region'] = data_4['Discharge_region'].str.replace("D  SINGAPORE", "SINGAPORE").str.strip()

#save in different spreadsheet
data = pd.concat([data_1, data_2,data_3,data_4], ignore_index = True,axis = 0)
data.drop_duplicates(inplace=True)

data.to_excel(path+'\CPPWest'+ re.sub(r'[^0-9]','',datetime.datetime.now().strftime("%d%m%Y")) + '.xlsx',index=False)
print('Done')


# In[ ]:




