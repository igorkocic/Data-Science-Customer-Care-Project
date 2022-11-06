#first
import pandas as pd
import re
import seaborn as sns
import matplotlib.pyplot as plt

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
sheets_dict = pd.read_excel('ABC calls 1.xlsx', sheet_name = None)

all_sheets = []
for name, sheet in sheets_dict.items():
    sheet['sheet'] = name
    sheet = sheet.rename(columns = lambda x: x.split('\n')[-1])
    all_sheets.append(sheet)

full_table = pd.concat(all_sheets)
full_table.reset_index(inplace = True)
full_table.rename(columns = {'sheet': 'day'}, inplace = True)
list_segment = []
for text in full_table['category']:
    found = re.findall(r"[S]\d{1}\s\w*", text)
    if found:
        list_segment.append(found)

my_string = str(list_segment).replace('[','').replace(']','').replace("'",'').replace(' ','')
list_of_string = my_string.split(',')
list_of_string1 = [x[2:] for x in list_of_string]
full_table['segment'] = list_of_string1
full_table['segment'].replace("LOST", 'LOST AND FOUND', inplace=True)

# Weekly NDA % by segment and global
attended = full_table.groupby('segment')['attended_number', 'received_number'].sum()
attended['%'] = 100 * (round(attended['attended_number'] / attended['received_number'],2))
# Weekly NDA % by segment
print(attended)
#Weekly NDA global percentage
nda_global = 100 * (attended['attended_number'].sum() / attended['received_number'].sum())
print("Global weekly percentage of answered is {}%".format(round(nda_global,2)))




# Weekly NDS% (<30 sec) by segment and global
full_table['attended_under_30_sec'] = full_table['attended_length_0_10'] + full_table['attended_length_10_20'] + full_table['attended_length_20_30']
attended_30sec = full_table[['segment', 'attended_under_30_sec', 'received_number']]
group2 = attended_30sec.groupby('segment')[['attended_under_30_sec', 'received_number']].sum()
# Weekly NDS% by segment
group2['%']= 100 * (round(group2['attended_under_30_sec'] / group2['received_number'],2))
print(group2)
# Weekly NDS global
nds_global = 100 * (attended_30sec['attended_under_30_sec'].sum() / attended_30sec['received_number'].sum())
print("Global weekly percentage of answered calls under 30 seconds of waiting is {}%".format(round(nds_global,2)))
print("According to the most common service target agents need to answer 80% of calls within 30 seconds.\nSince our global is slightly over 80% we can conclude that TARGET is ACHIEVED.")
print("Particular and lost and found segments need to improve their answered calls under 30 seconds")



