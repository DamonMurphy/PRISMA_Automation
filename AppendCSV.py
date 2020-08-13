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


