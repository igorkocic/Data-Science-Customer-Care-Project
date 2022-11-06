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


grouping_received = full_table.pivot_table(values = 'received_number', index = 'segment', columns = 'day')
grouping_attended = full_table.pivot_table(values = 'attended_number', index = 'segment', columns = 'day')
grouping_together = full_table.pivot_table(values = ['attended_number', 'received_number'], columns = 'day')
grouping_together_T = grouping_together.transpose()
grouping_together_T.reset_index(inplace = True)

# cats = [ 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
# df['Day of Week'] = df['Day of Week'].astype('category', categories=cats, ordered=True)

weekday = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY']

titles = list(grouping_together.columns)
titles[0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6] = titles[1], titles[5], titles[6], titles[4], titles[0], titles[2], titles[3]
grouping_together = grouping_together[titles]
print(grouping_together)


r=sns.barplot(x='day', y='received_number', hue='segment', data=full_table)
plt.ylabel('Number of received calls')
plt.xlabel('Weekday')
plt.title('Intra week evolution by segment - Number of received calls')
plt.show()

a=sns.barplot(x='day', y='attended_number', hue='segment', data=full_table)
plt.ylabel('Number of attended calls')
plt.xlabel('Weekday')
plt.title('Intra week evolution by segment - Number of attended calls')
plt.show()


received_list = grouping_together_T['received_number'].tolist()
attended_list = grouping_together_T['attended_number'].tolist()
day_list = grouping_together_T['day'].tolist()

plt.plot(day_list, received_list, label = 'received_number')
plt.plot(day_list, attended_list, label = 'attended_number')
plt.title('Received vs. Attended calls')
plt.ylabel('Number of calls')
plt.xlabel('Weekday')
plt.legend()
plt.show()
