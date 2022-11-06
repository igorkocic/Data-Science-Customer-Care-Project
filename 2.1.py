import pandas as pd
import numpy as np
import re
from datetime import timedelta
import matplotlib.pyplot as plt
import seaborn as sns
import os
print(os.getcwd())
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
full_table = pd.read_excel('ABC calls 2.xlsx')
pricing = pd.read_excel("price list.xlsx")

#Using regex in order to extract different segments from category column
list_segment = []
for text in full_table['category']:
    found = re.findall(r"[S]\d{1}\s\w*", text)
    if found:
        list_segment.append(found)

my_string = str(list_segment).replace('[','').replace(']','').replace("'",'')
list_of_string = my_string.split(',')
#list_of_string1 = [x[2:] for x in list_of_string]
full_table['segment'] = list_of_string

## S0, S1, S2, S4 segments cost estimation
grouped = pd.DataFrame(full_table.groupby('segment')['attended_number'].sum()).reset_index()
grouped['segment'].replace(" S4 LOST", "S4 LOST AND FOUND", inplace = True)
grouped = grouped.drop(grouped.index[[3,5]])
grouped1 = pd.concat([grouped, grouped], ignore_index=True)
grouped1['country'] = ['USA', 'USA', 'USA', 'USA', 'Canada', 'Canada', 'Canada', 'Canada']
grouped2 = pd.DataFrame(grouped1.groupby(['country','segment'])['attended_number'].sum()).reset_index()
list1 =  [float(pricing.loc[pricing['Segment'] == 'S0', 'Canada']), float(pricing.loc[pricing['Segment'] == 'S1', 'Canada']), float(pricing.loc[pricing['Segment'] == 'S2', 'Canada']), float(pricing.loc[pricing['Segment'] == 'S4', 'Canada']), float(pricing.loc[pricing['Segment'] == 'S0', 'USA']), float(pricing.loc[pricing['Segment'] == 'S1', 'USA']), float(pricing.loc[pricing['Segment'] == 'S2', 'USA']), float(pricing.loc[pricing['Segment'] == 'S4', 'USA'])]
grouped2['price'] = list1

def f(x):
    if x['country'] == 'Canada': return 0.45
    else: return 0.55
grouped2['%_of_calls_country'] = grouped2.apply(f, axis = 1)


grouped2['total_cost'] = grouped2['attended_number'] * grouped2['price'] * grouped2['%_of_calls_country']
print(grouped2)
print("Estimated total cost for S0, S1, S2, S4 segments taking into account the contribution of attended_calls by country and different prices is {}".format(round(grouped2['total_cost'].sum(),2)), "€")

#Data visualization
r = sns.catplot(data = grouped2, x='segment', y='total_cost', hue='country', kind='bar')
plt.xlabel('Segment')
plt.ylabel('Total_cost')
plt.title('Total cost by segments and countries')
plt.show()

# S3 cost estimation by country and LoB
grouped_S3 = pd.DataFrame(full_table.groupby('segment')['attended_number'].sum()).reset_index()
grouped_S3 = grouped_S3.drop(grouped_S3.index[[0,1,2,4,5]])
new_row = {'segment' : 'S3 SALES', 'attended_number' : 38751}
new_row1 = {'segment' : 'S3 SALES', 'attended_number' : 38751}
new_row2 = {'segment' : 'S3 SALES', 'attended_number' : 38751}
grouped_S3 = grouped_S3.append([new_row, new_row1, new_row2], ignore_index=True)
grouped_S3['LoB'] = ['Sales', 'Sales', 'Loyalty', 'Loyalty']
grouped_S3['country'] = ['USA', 'USA', 'Canada', 'Canada']
def s(z):
    if z['country'] == 'Canada': return 0.3
    else: return 0.70
grouped_S3['%_of_calls_country'] = grouped_S3.apply(s, axis = 1)
list2 = [float(pricing.loc[(pricing['Segment'] == 'S3') & (pricing['LoB'] == 'Sales'), 'USA']), float(pricing.loc[(pricing['Segment'] == 'S3') & (pricing['LoB'] == 'Loyalty'), 'USA']), float(pricing.loc[(pricing['Segment'] == 'S3') & (pricing['LoB'] == 'Sales'), 'Canada']), float(pricing.loc[(pricing['Segment'] == 'S3') & (pricing['LoB'] == 'Loyalty'), 'Canada'])]
grouped_S3['price'] = list2
grouped_S3['total_cost'] = grouped_S3['attended_number'] * grouped_S3['price'] * grouped_S3['%_of_calls_country']
print(grouped_S3)
print("Estimated total cost for S3 segment taking into account different prices by country is {}".format(round(grouped_S3['total_cost'].sum(),2)),"€")
print("Estimated grand total monthly cost of the service is {}".format(round(grouped2['total_cost'].sum() + grouped_S3['total_cost'].sum(),2)), "€")
