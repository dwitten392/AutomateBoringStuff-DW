#! python3

import re
import random
import pandas as pd
import datetime as dt
import math
import warnings

warnings.simplefilter(action='ignore')


line_check = re.compile(r'\s\d{4}\s') #the pattern to deterimine if a line should be imported

line_check1 = re.compile(r'\d{9}') #pattern to add a project name to the dict

name_dict = {}

while True:
    filename = input('Please enter text sales file name (or path if in different folder) or "A" to abort: ')
    if filename.lower() == 'a': exit() #end program if true
    elif filename[-4:].lower() != '.txt':
        filename = filename + '.txt'
    try:
        with open(filename) as f:
            switch = 0
            lines = []
            lines1 = [] 
            for line in f:
                if (line_check.search(line[10:16])) != None and (switch == 0): #imports sales file lines
                    lines.append(line.rstrip('\n'))

                elif (line_check1.search(line[4:13])) != None and (switch == 0): #import project name dictionary
                    name_dict[line[4:13]] = line[15:40].strip()

                elif 'Sales   Transfer' in line: switch = 1 #flips the switch to import transfers

                elif (line_check.search(line[10:16])) and (switch == 1): #import transfers
                    lines1.append(line.rstrip('\n'))
            f.close()
            break                  

    except: print('Invalid name; copy or type the whole path or put the text file in the same file as the working directory (program file)')

#SALES FILE (NOT TRANSFERS)
column_list = ['Project_Number', 'Project_Name', 'BUID', 'Business_Unit', 'Plat_Phs/Blk/Lot', 'Buyer_Name',
               'Begin_Date', 'Begin_Amount', 'Approve_Date', 'Approve_Amount',
               'Bustout_Date', 'Bustout_Amount', 'Close_Date', 'Close_Amount', 'Unique_ID']

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
        project_number.append(project_number[i-1]) #if the project number isn't listed, take from the previous pn
    else:
        project_number.append(lines[i][1:10]) #extract the project number

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

#creating project name using the dictionary
project_name = [name_dict[i] for i in project_number]

#defining business unit list
business_unit = [(int(project_number[i]) + int(buid[i])) for i in range(len(buid))]

#create unique ID
unique_id = [buyer_name[i] + str(business_unit[i]) for i in range(len(business_unit))]
            
list_data = [project_number, project_name, buid, business_unit, plat, buyer_name, begin_date, begin_amount, approve_date,
        approve_amount, bustout_date, bustout_amount, close_date, close_amount, unique_id]

#taking data to DF for more modifications
sales_df = pd.DataFrame(list_data).transpose()
sales_df.columns = column_list

#TRANSFERS
column_list1 = ['Project_Number', 'Project_Name', 'BUID', 'Business_Unit', 'Plat_Phs/Blk/Lot', 'Buyer_Name',
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

project_name1 = [name_dict[i] for i in project_number1]

business_unit1 = [(int(project_number1[i]) + int(buid1[i])) for i in range(len(buid1))]

list_data1 = [project_number1, project_name1, buid1, business_unit1, plat1, buyer_name1, to_buid1, to_plat1, sales_date1, transfer_date1]

transfer_df = pd.DataFrame(list_data1).transpose()
transfer_df.columns = column_list1

#converting close date to datetime column
sales_df.Begin_Date = pd.to_datetime(sales_df.Begin_Date, format = '%m/%d/%y', errors = 'coerce')

sales_df.Approve_Date = pd.to_datetime(sales_df.Approve_Date, format = '%m/%d/%y', errors = 'coerce')

sales_df.Bustout_Date = pd.to_datetime(sales_df.Bustout_Date, format = '%m/%d/%y', errors = 'coerce')

#Getting cutoff date from user
while True:
    date = input('What is your filter date. Format M-D-Y): ')
    
    try:   
        count = len(sales_df.loc[sales_df.Close_Date != ""].loc[sales_df[["Begin_Date", "Approve_Date"]].max(axis=1, skipna = True) >= pd.to_datetime(date)])
        count1 = len(sales_df.loc[sales_df.Close_Date != ""])
        break
    except:
        print('That is not a valid date, please try again.')
        continue

#telling the user the population number and gathering the number of samples
while True:
    
    print('There were ' + str(count1) + ' total sales closed. ' + str(count) + ' with the date restrictions.')
    pop_count = input('How many samples would you like to select? (0 for no selections): ')

    if pop_count == '0':
        break
    assert ((int(pop_count) <= 40) and (int(pop_count) > 0)) #raises an error if not true
        
    try:
        sales_sample_index = random.sample(list(sales_df.loc[sales_df.Close_Date != ""].loc[sales_df[["Begin_Date", "Approve_Date"]].max(axis=1, skipna = True) >= pd.to_datetime(date)].index), int(pop_count))
    
        sales_sample_df = sales_df.iloc[sales_sample_index, :]

        sales_sample_df.Begin_Date = sales_sample_df['Begin_Date'].dt.strftime('%m/%d/%Y')

        sales_sample_df.Approve_Date = sales_sample_df['Approve_Date'].dt.strftime('%m/%d/%Y')

    except:                 
        print('Please enter an integer between 0 and 40.')
        continue
    break

#sorting DF
sales_df = sales_df.sort_values(['Business_Unit', 'Begin_Date', 'Approve_Date'], ascending = (True, False, True))

#building out cancellation days list for DataFrame (probably a better way to do this)
cancellation_days = []
bustout_begin = 0
bustout_approve = 0

for i in range(len(sales_df)):
    
    bustout_begin = (sales_df.Bustout_Date[i] - sales_df.Begin_Date[i]).days
    if math.isnan(bustout_begin):
        bustout_begin = 0

    bustout_approve = (sales_df.Bustout_Date[i] - sales_df.Approve_Date[i]).days
    if math.isnan(bustout_approve):
        bustout_approve = 0
    
    cancellation_days.append(max(bustout_begin, bustout_approve))

sales_df["Contract_Cancel_Days"] = cancellation_days

#building out New Contract Days

sales_df["New_Contract_Days"] = ""

for i in range(len(sales_df.BUID)-1):
    if sales_df.Business_Unit.iloc[i] == sales_df.Business_Unit.iloc[i+1]:
        sales_df["New_Contract_Days"].iloc[i] = (sales_df["Approve_Date"].iloc[i+1] - sales_df["Bustout_Date"].iloc[i]).days

#telling the user the cancelation population number and gathering the number of samples
while True:

    print('There were ' + str(len(sales_df.loc[sales_df.Contract_Cancel_Days >= 30])) + ' cancellations with 30 days or more since the Begin/Approve date.')
    cancel_count = input('How many cancellation samples would you like to select? (0 for no selections): ')

    try:
        if cancel_count == '0':
            break
        assert (int(cancel_count) <= 40) and (int(cancel_count) > 0) #raises an error if not true
        
        cancel_sample_index = random.sample(list(sales_df[sales_df.Contract_Cancel_Days >= 30].index), int(cancel_count))

        cancel_sample_df = sales_df.iloc[cancel_sample_index, :]

        #appearances of DT columns

        cancel_sample_df.Begin_Date = cancel_sample_df['Begin_Date'].dt.strftime('%m/%d/%Y')

        cancel_sample_df.Approve_Date = cancel_sample_df['Approve_Date'].dt.strftime('%m/%d/%Y')

        cancel_sample_df.Bustout_Date = cancel_sample_df['Bustout_Date'].dt.strftime('%m/%d/%Y')

    except:
        print('Please enter an integer between 0 and 40.')
        continue
    break

#appearance of datetime columns
sales_df.Begin_Date = sales_df['Begin_Date'].dt.strftime('%m/%d/%Y')

sales_df.Approve_Date = sales_df['Approve_Date'].dt.strftime('%m/%d/%Y')

sales_df.Bustout_Date = sales_df['Bustout_Date'].dt.strftime('%m/%d/%Y')

#adding 1 #human indexing
sales_df.index = range(1, len(sales_df)+1) 
transfer_df.index = range(1, len(transfer_df)+1)
try: sales_sample_df.index = range(1, len(sales_sample_df)+1)
except: pass
try: cancel_sample_df.index = range(1, len(cancel_sample_df)+1)
except:pass
                
with pd.ExcelWriter(filename.split('.')[0] + '.xlsx') as writer:  # write to an excel sheet

    sales_df.to_excel(writer, sheet_name='Sales')

    transfer_df.to_excel(writer, sheet_name='Transfers')

    try: sales_sample_df.to_excel(writer, sheet_name='Sales Sample')
    except: pass

    try: cancel_sample_df.to_excel(writer, sheet_name='Cancel Sample')
    except: pass

print('Your excel file should be saved in your Python Scripts folder. Enjoy!')

