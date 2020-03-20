#! python3

import re
import os
import pandas as pd
from openpyxl import Workbook

line_check = re.compile(r'\d{4}\s\d') #the pattern to deterimine if a line should be imported
line_check1 = re.compile(r'\d{4}\s') #second pattern (a different method might be better)

while True:
    print('Please enter sales file name path (text) or "A" to abort')
    filename = input()
    if filename.lower() == 'a': exit() #end program if true        
    try:
        with open(filename) as f:
            switch = 0
            lines = []
            lines1 = [] #created to capture the transfers, will work on this later
            for line in f:
                if (line_check.search(line)) != None and (switch == 0):
                    lines.append(line.rstrip('\n'))  

                elif 'Transfer' in line: switch = 1 #flips the switch to import transfers

                elif (line_check1.search(line)) and (switch ==1):
                    lines1.append(line.rstrip('\n'))
            f.close()
            break                  

    except: print('Invalid file name; make sure you put the whole path or put the text file in the same file as the working directory (program file)')
    

#SALES FILE (NOT TRANSFERS)
column_list = ['Project_Number', 'BUID', 'Business_Unit', 'Plat_Phs/Blk/Lot', 'Buyer_Name',
               'Begin_Date', 'Begin_Amount', 'Approve_Date', 'Approve_Amount',
               'Bustout_Date', 'Bustout_Amount', 'Close_Date', 'Close_Amount']

project_number = [] #initializing empty lists to append to later
buid = []
plat = []
buyer_name = []
begin_date = []
begin_amount = []
approve_date = []
approve_amount = []
bustout_date = []
bustout_amount = []
close_date = []
close_amount = []

for i in range(len(lines)):
    if '  ' in lines[i][1:3]: # this part is to extract the project number
        project_number.append(project_number[i-1])
    else:
        project_number.append(lines[i][1:10])

    buid.append(lines[i][11:15]) #this part is to extract the BUID

    plat.append(lines[i][16:27].strip()) #extract the plat

    buyer_name.append(lines[i][40:65].strip()) #extract the buyer name

    begin_date.append(lines[i][65:73].strip()) #extract the begin date

    begin_amount.append(lines[i][79:91].strip()) #extract the begin amount

    approve_date.append(lines[i][91:100].strip()) #extract the approve date

    approve_amount.append(lines[i][105:118].strip()) #extract the approve amount

    bustout_date.append(lines[i][118:127].strip()) #extract the bustout date

    bustout_amount.append(lines[i][132:143].strip()) #extract the bustout amount

    close_date.append(lines[i][143:152].strip()) #extract the close date

    close_amount.append(lines[i][157:169].strip()) #extract the close date

#defining business unit list
business_unit = [(int(project_number[i]) + int(buid[i])) for i in range(len(buid))]
            
list_data = [project_number, buid, business_unit, plat, buyer_name, begin_date, begin_amount, approve_date,
        approve_amount, bustout_date, bustout_amount, close_date, close_amount]

#taking data to DF for more modifications
sales_df = pd.DataFrame(list_data).transpose()
sales_df.columns = column_list
sales_df.index = range(1, len(sales_df)+1) #adding 1 #human indexing

#TRANSFERS
column_list1 = ['Project_Number', 'BUID', 'Plat_Phs/Blk/Lot', 'Buyer_Name',
               'To_BUID', 'To_Plat_Phs/Blk/Lot', 'Sales_Date', 'Transfer_Date']

project_number1 = [] #initializing empty lists to append to later
buid1 = []
plat1 = []
buyer_name1 = []
to_buid1 = []
to_plat1 = []
sales_date1 = []
transfer_date1 = []

for i in range(len(lines1)):
    project_number1.append(lines1[i][1:10]) #append project number

    buid1.append(lines1[i][11:15]) #this part is to extract the BUID

    plat1.append(lines1[i][16:27].strip()) #extract the plat

    buyer_name1.append(lines1[i][40:65].strip()) #extract the buyer name

    to_buid1.append(lines1[i][65:69]) #this part is to extract the BUID

    to_plat1.append(lines1[i][70:80].strip()) #extract the plat

    sales_date1.append(lines1[i][93:102].strip()) #extract the sales date

    transfer_date1.append(lines1[i][103:112].strip()) #extra the transfer date

list_data1 = [project_number1, buid1, plat1, buyer_name1, to_buid1, to_plat1, sales_date1, transfer_date1]

transfer_df = pd.DataFrame(list_data1).transpose()
transfer_df.columns = column_list1
transfer_df.index = range(1, len(transfer_df)+1)

with pd.ExcelWriter(filename.split('.')[0] + '.xlsx') as writer:  # write to an excel sheet

    sales_df.to_excel(writer, sheet_name='Sales')

    transfer_df.to_excel(writer, sheet_name='Transfers')


