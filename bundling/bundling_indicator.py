import pandas as pd
import datetime

print('modules are imported')

# Load data
bundling81_data = pd.read_excel('bundling81.xlsx',
	                         sep=';', decimal=',',
	                         encoding='ANSI',
	                         skiprows=3, parse_dates=['Дата последнего отбора'])

print('bundling81_data is loaded')

# dont work auto parse dates(changed it manually for now)
pd.to_datetime(bundling81_data['Дата последнего отбора'], dayfirst = True, format = '%d/%m/%Y %I:%M')

# bundling81_data.to_excel('bundling81_mod.xlsx')

tasks81_data = pd.read_excel('tasks81.xlsx',
	                         sep=';', decimal=',',
	                         encoding='ANSI')

print('tasks81_data is loaded')

pd.to_datetime(tasks81_data['дата план отгр'], dayfirst = True, format = '%d/%m/%Y %I:%M')

# tasks81_data.to_excel('tasks81_mod.xlsx')

# Filter data
bundling81_data = bundling81_data.loc[bundling81_data['№ ЗнО (1С УПП)'].isin(tasks81_data['зно'])]

# Merge date and time of planning shippments
bundling81_data = pd.merge(bundling81_data, tasks81_data, left_on=['№ ЗнО (1С УПП)'], right_on=['зно'], how='left')

bundling81_data['intime'] = bundling81_data['дата план отгр'] > bundling81_data['Дата последнего отбора']

def calc_quant_intime(row):
	""""""
	if row['intime'] == True:
		row = row['Отобрано/ отгружено, шт']
	else:
		row = 0
	return row
		

bundling81_data['quantity intime'] = bundling81_data.apply(calc_quant_intime, axis='columns')
bundling81_data['percent of execution'] = bundling81_data['quantity intime'] / bundling81_data['Заказано, шт'] * 100

percent = bundling81_data['percent of execution'].mean()
rows = bundling81_data['№ ЗнО (1С УПП)'].count()

bundling81_data.to_excel('bundling81_mod.xlsx')
print(rows, percent)