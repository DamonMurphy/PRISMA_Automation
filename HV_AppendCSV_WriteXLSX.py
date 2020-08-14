# -*- coding: utf-8 -*-
"""
Created on Thu Aug 13 14:15:23 2020

@author: murphd26
"""

from dateutil.relativedelta import relativedelta
import pandas as pd
from datetime import date, datetime
import calendar
import os, sys
from os import listdir
from os.path import isfile
import shutil
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment

prisma_period = (date.today() + relativedelta(months=-1)).replace(day=1)
prisma_period = pd.Timestamp(prisma_period)

# DETERMINE FOLDER LOCATION
prisma_month = prisma_period.month
prisma_year = prisma_period.year
prisma_month_abbr = calendar.month_abbr[prisma_month]
prisma_month_name = calendar.month_name[prisma_month]


directory_6270 = r'\\Ohcolnas0250\PSHP\PRISMA_Files'

os.chdir(directory_6270)
os.chdir(str(prisma_year))


year_path = os.getcwd()
dir_list = [name for name in os.listdir(year_path) if os.path.isdir(os.path.join(year_path, name))]

dir_name = [dir1 for dir1 in dir_list if (prisma_month_abbr in dir1 or prisma_month_name in dir1)]

if len(dir_name) != 1:
    print('\n\nThere are' + str(len(dir_name)) + 'names in the folder list for the reporting month.')
    print('Check',year_path,'for correct folder then continue when fixed.\n\n')
    print('Quiting...\n\n')
    os.exit(0)

dir_name = dir_name[0]

os.chdir(dir_name)

month_path = os.getcwd()
dir_list = [name for name in os.listdir(month_path) if os.path.isdir(os.path.join(month_path, name))]

dir_name = [dir1 for dir1 in dir_list if '6270' in dir1]

if len(dir_name) != 1:
    print('\n\nThere are' + str(len(dir_name)) + 'names in the folder list for the 6270 files.')
    print('Check',month_path,'for correct folder then continue when fixed.\n\n')
    print('Quiting...\n\n')
    os.exit(0)

dir_name = dir_name[0]

os.chdir(dir_name)

path = os.getcwd()


fileList = [name for name in os.listdir(path) if 'CSV' in str.upper(name)]


file1 = prisma_month_name + '6270.csv'

with open(file1, 'w') as outfile:
    for fname in fileList:
        if isfile(fname):
            with open(fname) as infile:
                outfile.write(infile.read())

outfile.close()


shutil.move(path + '\\' + file1,month_path + '\\' + file1)


os.chdir(month_path)

df_HV_checks = pd.read_csv(file1,sep=';', lineterminator='\r')

df_HV_checks = df_HV_checks[df_HV_checks['username'] == 'PRODUSER']

# RESET INDEX
df_HV_checks = df_HV_checks.reset_index(drop=True)

df_HV_checks = df_HV_checks[['jobid','startdate','result', 'username', 'jobname','nofprinteda4bw']]

# RENAME COLUMNS
df_HV_checks.rename(columns={'jobid':'JobID','startdate':'StartDate',\
        'result':'Status','username':'UserName','jobname':'JobName',\
            'nofprinteda4bw':'Images'},inplace=True)


# CHECK IF Status OTHER THAN 'Done' IS PRESENT
status_entries = df_HV_checks['Status'].drop_duplicates(keep='first').reset_index(drop=True)
if (status_entries != 'Done').any():
    print('\n\nThere are entries other than \'Done\' in this data set.')
    print('Check and see how to handle.\n')
    print('Other \'Status\' values:')
    print(status_entries[(status_entries != 'Done')])
    print('\n\nQuiting....\n\n')
    sys.exit(0)




# REMOVE BLANK SPACES FROM JobName
df_HV_checks['JobName'] = df_HV_checks['JobName'].str.strip().str.replace(' ','')


# FORMAT StartDate AS DATE
df_HV_checks['StartDate'] = pd.to_datetime(df_HV_checks['StartDate'],infer_datetime_format=True)


# FORMAT Images AS INT
df_HV_checks['Images'] = df_HV_checks['Images'].astype('int')


# WRITE NEW EXCEL FILE
file1 = prisma_month_name + '6270_Checks.xlsx'
ws1 = prisma_month_abbr + str(prisma_year)

print()
print()
print('writing data to xlsx file...')
print()
print()

HV_checksWB = Workbook()
HV_checksWB.title = file1

HV_checksWS = HV_checksWB.active
HV_checksWS.title = ws1


colNames = list(df_HV_checks.columns)

i = 65
for colName in colNames:
    #print([chr(i)+'1'],'\t',colName)
    HV_checksWS[chr(i)+'1'] = colName
    HV_checksWS[chr(i)+'1'].font = Font(bold = True)
    if ('Date' in colName or 'Images' in colName):
        HV_checksWS[chr(i)+'1'].alignment = Alignment(horizontal='right')
    i+=1

print('\nHeaders finished\n\n')

NumberOfRecords = len(df_HV_checks)

print('\nNumberOfRecords:',NumberOfRecords)
print('\n')


for j in range(NumberOfRecords):
    HV_checksWS['B'+str(j+2)] = df_HV_checks.index[j]
    HV_checksWS['B'+str(j+2)].number_format = 'M/D/YYYY'


for i in range(len(df_HV_checks.columns)):
    if i % 5 == 0:
            print('column',i,'out of',len(df_HV_checks.columns),'columns')
    for j in range(NumberOfRecords):
        HV_checksWS[chr(i+66)+str(j+2)] = df_HV_checks.iloc[j,i]


print()
print()
print('closing Final xlsx...')
print()
print()



print('saving file...')
HV_checksWB.save(file1)
print('file saved')
print('closing file...')
HV_checksWB.close()
print('file closed')
del HV_checksWB
print('outputWorkbook deleted')



print(df_HV_checks['Images'].astype('int').sum())



sys.exit(0)

