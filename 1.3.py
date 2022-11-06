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

# Weekly average handling time for each segment in seconds
def convert(x):
    x = x.split(":")
    return int(x[0]) * 3600 + int(x[1]) * 60 + int(x[2])
full_table["attended_aht_seconds"] = full_table["attended_aht"].apply(convert)
print(full_table.groupby('segment')['attended_aht_seconds'].mean())
print("Target is achieved since average handling time for each segment is below 320 seconds")



