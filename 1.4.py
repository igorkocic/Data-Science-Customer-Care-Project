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


group = full_table.groupby(['segment', 'day'])[['abandoned_number', 'overflow_total']].sum()
print(group)
print(full_table['abandoned_number'].mean())
# Let's extract segments and days when abandoned number is over the mean.
over_mean = group[group['abandoned_number']>2.46]
print("Further analysis needed in order to discover why there are deviations on certain day and in certain segments.\nEspecially in segment 2 where we have extreme values for every day of the week.")
over_mean_df = over_mean.sort_values(by = ['abandoned_number'], ascending = False)
print(over_mean_df)
print(type(over_mean_df))
# With regards to overflowed calls, there is no need to improve processes since we have only one overflowed call for whole week.



