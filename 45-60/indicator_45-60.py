import pandas as pd
import datetime
import calendar
import math
import matplotlib.pyplot as plt
import seaborn as sns

print('modules are imported')

month = 4
year = 2020
holidays = [1, 2, 3, 6, 7, 8]
outsiders = ['Компания НБК ООО', 'СК-АВТО, ООО', 'Портал']
target = None

today = datetime.datetime.today()
# today = datetime.date(2020, 4, 9)

# today_ftd = '{}.{}.{}'.format(today.day, today.month, today.year)
today_ftd = today.strftime('%d.%m.%Y')

cal_obj = calendar.Calendar(0)
month_arr = cal_obj.monthdayscalendar(year, month)

work_days = []

for week_index, week in enumerate(month_arr):
	for day_index, day in enumerate(week[:5]):
		if month_arr[week_index][day_index] not in (holidays + [0,]):
			work_days.append(month_arr[week_index][day_index])

work_days = work_days[:math.ceil(len(work_days)/2)]

if today.day in work_days:
	target = 95
else:
	target = 99

#read csv
# reminder = pd.read_csv('reminder.csv',
# 						sep=';', decimal=',',
# 						encoding='ANSI')

#read xlsx
reminder_data = pd.read_excel('reminder.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',
						skiprows=5)
						#index_col='ЗаказПокупателяКонтрагент')

intasks_data = pd.read_excel('intasks.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',
						skiprows=5)

orders_data = pd.read_excel('orders.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',
						skiprows=5)												



reminder = round(reminder_data.loc[(~reminder_data['ЗаказПокупателяКонтрагент'].isin(outsiders)),
					'СуммаВзаиморасчетовОстаток'].sum()/1000, 2)

intasks = round(intasks_data.loc[(~intasks_data['Контрагент'].isin(outsiders)),
					'КЦсНДС'].sum()/1000, 2)

orders = round(orders_data.loc[(~orders_data['Контрагент'].isin(outsiders)),
					'Поле1'].sum()/1000, 2)

done = orders - reminder

data = [[orders,
		reminder,
		done,
		intasks,
		round((orders * (target/100)) - done, 2),
		target,
		round(done / orders * 100, 2),
		round((done - intasks) / orders * 100, 2)
		]]

col_names = ['Сформировано', 'Остаток ЗП', 'Выдано+Отгр',
			 'В ЗнО', 'Нужно еще посадить в ЗнО', 'Цель, %',
			 '% отр с ЗнО', '% отр без ЗнО']

perf_today = pd.DataFrame(data, columns = col_names, index=[today_ftd])

perf_overall = pd.read_excel('performance.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',
						index_col=0)

performance = pd.concat([perf_overall, perf_today])

# Delete last row in the DataFrame if last two indicies(dates) are equivalent
if performance.iloc[-1].name == performance.iloc[-2].name and performance.iloc[-1].name != None :
	performance = performance[:-1]

# performance.to_excel('performance.xlsx')

# Ploting and save data into .png file
plt.figure(figsize=(10, 5))
sns.set_style('whitegrid')
plt.title('Динамика исполнения показателя 45/60 дней')
sns.lineplot(data = performance['% отр без ЗнО'], label = '% исполнения')
sns.lineplot(data = performance['Цель, %'], label = 'Цель, %')
plt.xlabel('Дата')
plt.ylabel('%')
plt.legend()

plt.savefig('plot.png', bbox_inches='tight', orientation='landscape')
print('plot saved')

#from https://xlsxwriter.readthedocs.io/example_pandas_column_formats.html
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("performance.xlsx", engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
performance.to_excel(writer, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Add some cell formats.
format1 = workbook.add_format({'num_format': '# ##0'})
format2 = workbook.add_format({'num_format': '# ##0,00'})

# Set the column width and format.
worksheet.set_column('B:B', 15, format1)
worksheet.set_column('C:C', 10, format1)
worksheet.set_column('D:D', 13, format1)
worksheet.set_column('E:E', 10, format1)
worksheet.set_column('F:F', 26, format1)
worksheet.set_column('G:G', 8, format1)
worksheet.set_column('H:H', 11)
worksheet.set_column('I:I', 13)

# Insert plot in excel file
worksheet.insert_image('K1', 'plot.png')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

print('done')

# with pd.ExcelWriter('performance.xlsx', mode='a') as writer:
# 	performance.to_excel(writer, sheet_name='Sheet1')
