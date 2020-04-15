import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

print('modules are imported')


# Load shipped products data
shipped_data = pd.read_excel('fact_schedule.xlsx',
	                         sep=';', decimal=',',
	                         encoding='ANSI',
	                         skiprows=5)

print('shipped_data uploaded')

# Delete last empty row
shipped_data.drop(shipped_data.tail(1).index, inplace=True)

# Group by customer
shipped_data['Поле1'] = shipped_data['Поле1'] / 1000
shipped_by_customer = shipped_data.groupby('Грузополучатель').Поле1.sum()


print('shipped_data processed')

# shipped_by_customer.to_excel('fact_by_customer_mod.xlsx')



# Load schedule data
schedule_data = pd.read_excel('schedule.xlsx', sep=';',
	                          decimal=',', encoding='ANSI',
	                          skiprows=4)

print('schedule_data uploaded')

# Drop first and last empty rows
schedule_data.drop(schedule_data.head(1).index, inplace=True)
schedule_data.drop(schedule_data.tail(1).index, inplace=True)

schedule_data.drop(columns = schedule_data.columns[42:], axis = 1, inplace=True)

schedule_data = schedule_data.loc[schedule_data['Код заказа'] == 'КЦ']


# Specify last day to include in sum range
end_range = 31

schedule_data['Plan'] = schedule_data.loc[:, schedule_data.columns[11:end_range+11]].sum(axis=1)

schedule_data = schedule_data.groupby('Грузополучатель.1').Plan.sum()
# schedule_data.to_excel('schedule_by_customer_mod.xlsx')

print('schedule_data processed')
# schedule_data.to_excel('schedule_mod.xlsx')

report = pd.concat([schedule_data, shipped_by_customer], axis=1, sort=False)

report.rename(columns={'Поле1' : 'Fact'}, inplace=True)

# Replacing missing values with 0
report['Fact'].fillna(0, inplace=True)

report['Deviation'] = report['Fact'] - report['Plan']

def calc_execution(row):
	""""""
	if row['Plan'] > 0:
		row = row['Fact'] / row['Plan'] * 100
	else:
		row = 0
	return row
		

report['Execution'] = report.apply(calc_execution, axis='columns')

upper_limit = 105.4
lower_limit = 94.5
NDS = 1.2

def calc_noexecution(row):
	""""""
	if row['Execution'] > upper_limit:
		row = row['Execution'] - upper_limit
	elif 0 < row['Execution'] < lower_limit:
		row = lower_limit - row['Execution']
	elif (row['Plan'] == 0 and row['Fact'] > 0) or (row['Plan'] > 0 and row['Fact'] == 0):
		row = lower_limit
	else:
		row = 0
	return row				


report['Not_execution'] = report.apply(calc_noexecution, axis = 'columns')

def calc_sum_noex(row):
	""""""
	if row['Plan'] > 0:
		row = row['Plan'] * row['Not_execution'] / 100
	elif row['Fact'] > 0:
		row = row['Fact'] * row['Not_execution'] / 100
	else:
		row = 0
	return row		

report['Not_ex_sum'] = report.apply(calc_sum_noex, axis = 'columns')
report['Wtht_nds'] = report['Not_ex_sum'] / NDS

report.sort_values(by = 'Execution', inplace = True)

total_arr = []
columns_perc = ['Execution', 'Not_execution']

for col_name in report.columns:
	if col_name not in columns_perc:
		total_arr.append(report[col_name].sum())
	else:
		total_arr.append(report[col_name].mean())


total_arr_wo_nds = []

for col_name, arr_el in zip(report.columns, total_arr):
	if col_name not in columns_perc + ['Wtht_nds']:
		total_arr_wo_nds.append(arr_el / NDS)
	else:
		total_arr_wo_nds.append(arr_el)	

total_row = pd.DataFrame([total_arr], columns = report.columns, index = ['Total'])
total_row_wo_nds = pd.DataFrame([total_arr_wo_nds], columns = report.columns, index = ['Total_wo_nds'])

# columns_todivide = ['Plan', 'Fact', 'Deviation', 'Not_ex_sum']
# columns_wo_changes = ['Execution', 'Not_execution', 'Wtht_nds']


# total_row_wo_nds = (total_row_wo_nds[columns_todivide] / NDS) + total_row_wo_nds[columns_wo_changes]

report = report.append(total_row)
report = report.append(total_row_wo_nds)

report.index.name = 'Customer'
# report.to_excel('report.xlsx')


# Ploting and save data into .png file
plt.figure(figsize=(5, 10))
execution_plot = sns.barplot(x = report['Execution'],
	                         y = report.index)


for i, v in enumerate(report['Execution']):
    execution_plot.text(v + 1, i + .25, str(round(v)), color='black')


plt.savefig('plot.png', bbox_inches='tight', orientation='landscape')
print('barh saved')


# Handle cells format with XlsxWriter module
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("report.xlsx", engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
report.to_excel(writer, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Add some cell formats.
format1 = workbook.add_format({'num_format': '# ##0'})
format2 = workbook.add_format({'num_format': '# ##0,00'})

# Set the column width and format.
worksheet.set_column('A:A', 35, format1)
worksheet.set_column('B:B', 9, format1)
worksheet.set_column('C:C', 9, format1)
worksheet.set_column('D:D', 9, format1)
worksheet.set_column('E:E', 14, format1)
worksheet.set_column('F:F', 14, format1)
worksheet.set_column('G:G', 14, format1)
worksheet.set_column('H:H', 9, format1)

# Insert plot in excel file
worksheet.insert_image('J2', 'plot.png')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

print('done')