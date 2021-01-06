# EDI Version 0.1.1
import win32com.client
import re
import pandas as pd
import numpy as np
from datetime import datetime, date

with open('reconfig.txt', 'r') as file:
    outlook_folder_name = file.readline()

print(outlook_folder_name)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders.Item(outlook_folder_name)
messages = inbox.Items

# for m in messages:
#     print(m.subject)
#     print(m.CreationTime)
#     print(m.body)
#     print('='*80)

print('\nThere are {} emails in your mailbox\{}\n'.format(len(messages), outlook_folder_name))

mails = [np.NAN for i in range(len(messages))]
errors = ['missing' for i in range(len(messages))]
creation_dates = [np.NAN for i in range(len(messages))]
for i, message in enumerate(messages):
    mails[i] = message.subject
    #print(message.subject)
    try:
        creation_dates[i] = message.CreationTime.strftime('%Y-%m-%d')
    except:
        print(f'{mails[i]} - something wrong with date format (will be missing)')
    try:
        body_content = message.body
        #a = re.findall('Error Notes(?s)(.+)Data Model', body_content)[0]
        a = re.findall('(?s)Error Notes(.+)Data Model', body_content)[0]
        b = re.findall('.*', a)
        for j in b:
            if 'Error: PO:' in j: 
                errors[i] = j
                break
                
    except:
        print('Mail was not read properly: {}'.format(message.subject))

df = pd.DataFrame({'error':errors, 'mail':mails, 'receivedate':creation_dates})

df['error'] = df['error'].str.replace(' \r','')


def extract_numbers_from_text(x):
    error_msg = x
    #'po','poitem','product','variable'
    values_in_error_msg = []
    try:
        for i in re.finditer(r'\[ (\w+|\s|\d+|\d+\.\d+) \]', error_msg):
            values_in_error_msg.append(error_msg[i.span()[0]+2:i.span()[1]-2])
    except:
        pass
    if len(values_in_error_msg) == 0:
        values_in_error_msg = ['missing','missing','missing','missing']
    elif len(values_in_error_msg) < 4:
        for _ in range(4 - len(values_in_error_msg)):
            values_in_error_msg.append('')
    #print(len(values_in_error_msg), values_in_error_msg)
    return values_in_error_msg

df['po'] = df['error'].apply(lambda x: extract_numbers_from_text(x)[0])

df['item'] = df['error'].apply(lambda x: extract_numbers_from_text(x)[1])

df['product'] = df['error'].apply(lambda x: extract_numbers_from_text(x)[2])
df['variable'] = df['error'].apply(lambda x: extract_numbers_from_text(x)[-1])

df['error'] = df['error'].apply(lambda x: x[x.find('-')+2:] if x.find('-')!=-1 else x)

def error_type(x):
    error_types = {'Ship To Code:':'Ship To Code', 'Odd/Last Carton': 'Carton', 'Item Codes:': 'Material Missing',
                  'FCL Port of Discharge': 'Port Missing', 'Supplier Code :': 'Supplier Missing',
                  'missing':'NoError',
                  'Supplier Name': 'Supplier Name Missing',
                  'Invalid Funloc Code': 'Invalid Funloc Code',
                  'MAG': 'MAG missing',
                  ': 	Error: PO': 'Missing DTM', 
                  ':Error: PO': 'Missing DTM'}
    
    error_return = 'Other'
    for k, v in error_types.items():
        #print(k, v)
        if x.startswith(k):
            error_return = v
            
    return error_return

df['error_type'] = df['error'].apply(error_type)

df = df[['mail', 'error', 'receivedate', 'error_type', 'po', 'item', 'product', 'variable']]

#df.drop_duplicates(inplace=True)
print('===== Errors found =====:\n', df.pivot_table(values='variable', index='error_type', aggfunc=['count', pd.Series.nunique]))


df_sum = df.pivot_table(values='variable', index='error_type', aggfunc=['count', pd.Series.nunique])
df_sum.columns = ['count','unique variables']
df_sum.sort_values('unique variables', ascending=False, inplace=True)

try:
    df_carton = pd.DataFrame(df.loc[df.error_type=='Carton'].groupby(['error_type', 'product']).size())
    df_carton.columns = ['counts']
    df_carton.sort_values('counts', ascending=False, inplace=True)
except:
    df_carton = pd.DataFrame()

try:
    df_shiptocode = pd.DataFrame(df.loc[df.error_type=='Ship To Code'].groupby(['error_type', 'variable']).size())
    df_shiptocode.columns = ['counts']
    df_shiptocode.sort_values('counts', ascending=False, inplace=True)
except:
    df_shiptocode = pd.DataFrame()

try:    
    df_matmis = pd.DataFrame(df.loc[df.error_type=='Material Missing'].groupby(['error_type', 'variable']).size())
    df_matmis.columns = ['counts']
    df_matmis.sort_values('counts', ascending=False, inplace=True)
except:
    df_matmis = pd.DataFrame()
    
try:
    df_supmis = pd.DataFrame(df.loc[df.error_type=='Supplier Missing'].groupby(['error_type', 'variable']).size())
    df_supmis.columns = ['counts']
    df_supmis.sort_values('counts', ascending=False, inplace=True)
except:
    df_supmis = pd.DataFrame()

try:
    df_invalid_funloc = pd.DataFrame(df.loc[df.error_type=='Invalid Funloc Code'].groupby(['error_type', 'variable']).size())
    df_invalid_funloc.columns = ['counts']
    df_invalid_funloc.sort_values('counts', ascending=False, inplace=True)
except:
    df_invalid_funloc = pd.DataFrame()

try:
    df_supnamemis = pd.DataFrame(df.loc[df.error_type=='Supplier Name Missing'].groupby(['error_type', 'error']).size())
    df_supnamemis.columns = ['counts']
    df_supnamemis.sort_values('counts', ascending=False, inplace=True)
except:
    df_supnamemis = pd.DataFrame()



excel_writer = pd.ExcelWriter('edi errors - {}.xlsx'.format(datetime.now().strftime('%Y-%m-%d %H%M')))
df.to_excel(excel_writer, index=False, freeze_panes=(1,0), sheet_name='data')
df_sum.to_excel(excel_writer, freeze_panes=(1,0), sheet_name='summary')
if df_carton.shape[0] > 0:
    df_carton.to_excel(excel_writer, freeze_panes=(1,0), sheet_name='Carton')
if df_shiptocode.shape[0] > 0:
    df_shiptocode.to_excel(excel_writer, freeze_panes=(1,0), sheet_name='Ship To Code')
if df_matmis.shape[0] > 0:
    df_matmis.to_excel(excel_writer, freeze_panes=(1,0), sheet_name='Material Missing')
if df_supmis.shape[0] > 0:
    df_supmis.to_excel(excel_writer, freeze_panes=(1,0), sheet_name='Supplier Missing')
if df_supnamemis.shape[0] > 0:
    df_supnamemis.to_excel(excel_writer, freeze_panes=(1,0), sheet_name='Supplier Name Missing')
if df_invalid_funloc.shape[0] > 0:
    df_invalid_funloc.to_excel(excel_writer, freeze_panes=(1,0), sheet_name='Invalid Funloc')

excel_writer.save()

print('\nDone.')
