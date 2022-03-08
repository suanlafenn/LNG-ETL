#!/usr/bin/env python
# coding: utf-8

# In[13]:


import datetime
import re
import os
import win32com.client
import pandas as pd
import numpy as np
import tabula
def has_numbers(inputString):
     return any(char.isdigit() for char in inputString)

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox=outlook.GetDefaultFolder(6).Folders("freight")
Name =[]
messages=inbox.Items

for message in messages:
    if 'SIMPSON SPENCE YOUNG East of Suez(' in message.Subject:
        attachments = message.attachments
        for attachment in attachments:
            attachment.SaveAsFile(os.getcwd() + '\\' + attachment.FileName)
            Name.append(attachment.FileName)
Vessel_name =[]
status = []
Charterer = []
Cargo = []
Type = []
Loading_region=[]
Discharge_region =[]
load_date = []
Freight_rate = []
pdf_name =[]
for i in Name:
    print(i)
    path = i
    tabula.convert_into(path, "output.csv",output_format="csv", pages= '1',area=[200,0,780,460])
    df = pd.read_csv('output.csv',encoding='cp1252')
    df.dropna(how='all', axis=1, inplace=True)
    if len(df.columns) == 3:
        df = df.set_axis(['A', 'B', 'C'], axis=1)
        df.dropna(subset=(['B', 'C']),how='all', inplace = True)
        df = df.iloc[0:-1]
        df_list = df.values.tolist()        
        for i in df_list:
            if 'fxd' in i[0]:
                text = i[0][:i[0].index('fxd')]
            elif 'subs' in i[0]:
                text = i[0][:i[0].index('subs')]
            elif 'fld' in i[0]:
                text = i[0][:i[0].index('fld')] 
            Vessel_name.append(text.rstrip())

        for i in df_list:
            if 'fxd' in i[0]:
                text = 'fxd'
            elif 'subs' in i[0]:
                text = 'subs'
            elif 'fld' in i[0]:
                text = 'fld'
            status.append(text.rstrip())

        for i in df_list:
            if 'fxd' in i[0]:
                text = re.search(r'(?<=fxd).*(?=\d{2})', i[0])
            elif 'subs' in i[0]:
                text = re.search(r'(?<=subs).*(?=\d{2})', i[0])
            elif 'fld' in i[0]:
                text = re.search(r'(?<=fld).*(?=\d{2})', i[0])
            Charterer.append(text.group().strip())    

        for i in df_list:
            text = re.search(r'\d{2}', i[0]).group().strip()
            Cargo.append(text)   

        for i in df_list:
            text = re.search(r'(?<=\d{2}).*', i[0]).group()
            Type.append(text.split(None,1)[0])  
        for i in df_list:
            text = re.search(r'(?<=\d{2}).*', i[0]).group()
            Loading_region.append(text.split(None,1)[1]) 

        for i in df_list:
            if 'Chittagong-Sp end/Feb' in str(i[1]):
                i[1] = 'Chittagong-Sp 28/Feb'
            elif 'Eafr-Safr ely/Mar' in str(i[1]):
                i[1] = 'Chittagong-Sp 01/Mar'
            elif 'Eafr-Safr end/Feb' in str(i[1]):
                i[1] = 'Eafr-Safr 28/Feb'    
            elif 'Oz ely/Feb' in str(i[1]):
                i[1] = 'Oz 01/Feb' 
            elif 'Qatar end/Jan' in str(i[1]):
                i[1] = 'Qatar 31/Jan' 
            if has_numbers(str(i[1])) == True:
                text = re.search(r'.*(?=\d{2}/\w{3})', str(i[1])).group()
                text_2 = re.search(r'\d{2}/\w{3}', i[1]).group()
                text_3 = i[2]
            elif has_numbers(str(i[1])) == False:
                text = i[1]
                text_2 = re.search(r'\d{2}/\w{3}', i[2]).group()
                text_3 = re.search(r'(?<=\d{2}/\w{3}).*', i[2]).group()
            load_date.append(text_2) 
            Discharge_region.append(str(text).strip())
            Freight_rate.append(text_3)
            
    elif len(df.columns) == 4:
        df = df.set_axis(['A', 'B', 'C','D'], axis=1)
        df.dropna(subset=(['B', 'C','D']),how='all', inplace = True)
        df = df.iloc[0:-1]
        df_list = df.values.tolist()
        for i in df_list:
            if 'fxd' in i[0]:
                text = i[0][:i[0].index('fxd')]
            elif 'subs' in i[0]:
                text = i[0][:i[0].index('subs')]
            elif 'fld' in i[0]:
                text = i[0][:i[0].index('fld')] 
            Vessel_name.append(text.rstrip())

        for i in df_list:
            if 'fxd' in i[0]:
                text = 'fxd'
            elif 'subs' in i[0]:
                text = 'subs'
            elif 'fld' in i[0]:
                text = 'fld'
            status.append(text.rstrip())

        for i in df_list:
            if 'fxd' in i[0]:
                text = re.search(r'(?<=fxd).*(?=\d{2})', i[0])
            elif 'subs' in i[0]:
                text = re.search(r'(?<=subs).*(?=\d{2})', i[0])
            elif 'fld' in i[0]:
                text = re.search(r'(?<=fld).*(?=\d{2})', i[0])
            Charterer.append(text.group().strip())    

        for i in df_list:
            text = re.search(r'\d{2}', i[0]).group().strip()
            Cargo.append(text)   

        for i in df_list:
            text = re.search(r'(?<=\d{2}).*', i[0]).group()
            Type.append(text.split(None,1)[0])   

        for i in df_list:
            if 'Ums+TameMumbai+Sohar' in str(i[0]):
                i[0] = 'NCC Tabuk  subs OQ 35 Ums TameMumbai+Sohar'
            text = re.search(r'(?<=\d{2}).*', i[0]).group()
            Loading_region.append(text.split(None,1)[1])  

        for i in df_list:
            Discharge_region.append(i[1])
            load_date.append(i[2])
            Freight_rate.append(i[3])
    os.remove(path)
data = pd.DataFrame({'Vessel_name': Vessel_name,'Cargo': Cargo,'Type': Type,'Loading_region': Loading_region,
                      'Discharge_region': Discharge_region,'load_date': load_date,'Freight_rate': Freight_rate,
                     'Charterer': Charterer,'status': status})
data.drop_duplicates(inplace=True)
pathh = os.path.join(os.path.expanduser("~"), 'Desktop')
data.to_excel(pathh+'\SIMPSON_SPENCE_pdf'+ re.sub(r'[^0-9]','',datetime.datetime.now().strftime("%d%m%Y")) + '.xlsx',index=False)
os.remove("output.csv")
print('done')



# In[ ]:




