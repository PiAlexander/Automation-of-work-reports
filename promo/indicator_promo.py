import pandas as pd
import datetime
import matplotlib.pyplot as plt
import seaborn as sns

print('modules are imported')

outsiders = ['Компания НБК ООО', 'СК-АВТО, ООО', 'Портал']
target = 95

today = datetime.datetime.today()
# today = datetime.date(2020, 4, 9)

# today_ftd = '{}.{}.{}'.format(today.day, today.month, today.year)
today_ftd = today.strftime('%d.%m.%Y')

def calculate_sums(df_list):
	"""Adds new columns in DF"""

	output = []

	for df in df_list:
		df['contract'] = df['ДоговорКонтрагента'].map(lambda p: p[29:31])
		df['rep_pr'] = df['НоменклатураАртикул'].isin(rep_pr_data['ND'])
		df['wheels'] = df['НоменклатураАртикул'].isin(wheels_data['ND'])
		df['ord_intime'] = df['Контрагент'].isin(orders_data['Контрагент'])

		output.append(df.loc[((~df['Контрагент'].isin(outsiders)) &
							  (df['contract'] == 'ДО') &
							  (df['rep_pr'] == True) &
							  (df['wheels'] == False) &
							  (df['ord_intime'] == True)),
							  'СуммаВзаиморасчетовОстаток'].sum())
	return output

#read xlsx
# reminder_data = intasks_data = orders_data = rep_pr_data = wheels_data = None

# var_names = [reminder_data, intasks_data, orders_data, rep_pr_data, wheels_data]
# files = ['reminder_promo.xlsx', 'intasks_promo.xlsx', 'orders_promo.xlsx', 'reporting_products.xlsx',
# 		 'wheels.xlsx']

# for index, df in enumerate(var_names):
# 	df = pd.read_excel(files[index],
# 						sep=';', decimal=',',
# 						encoding='ANSI',)	 


reminder_data = pd.read_excel('reminder_promo.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',
						skiprows=5)

reminder_data.drop(reminder_data.tail(1).index, inplace=True)
reminder_data = reminder_data.rename(columns={'ЗаказПокупателяКонтрагент' : 'Контрагент',
							  				  'ЗаказПокупателяДоговорКонтрагента' : 'ДоговорКонтрагента'})

print('reminder_data uploaded')

						
intasks_data = pd.read_excel('intasks_promo.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',
						skiprows=5)

intasks_data.drop(intasks_data.tail(1).index, inplace=True)
intasks_data = intasks_data.rename(columns={'Номенклатура.Артикул' : 'НоменклатураАртикул',
											'КЦсНДС' : 'СуммаВзаиморасчетовОстаток'})

print('intasks_data uploaded')

orders_data = pd.read_excel('orders_promo.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',
						skiprows=5)

orders_data.drop(orders_data.tail(1).index, inplace=True)
orders_data = orders_data.rename(columns={'Поле1' : 'СуммаВзаиморасчетовОстаток'})

print('orders_data uploaded')


rep_pr_data = pd.read_excel('reporting_products.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',)

print('rep_pr_data uploaded')


wheels_data = pd.read_excel('wheels.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',)

print('wheels_data uploaded')


dataframes = [reminder_data, intasks_data, orders_data]

reminder, intasks, orders = calculate_sums(dataframes)

print('sums calculated')


done = orders - reminder

data = [[round(orders/1000, 2),
		round(reminder/1000, 2),
		round(done/1000, 2),
		round(intasks/1000, 2),
		round(((orders * (target/100)) - done)/1000, 2),
		target,
		round(done / orders * 100, 2),
		round((done - intasks) / orders * 100, 2)
		]]

col_names = ['Сформировано', 'Остаток ЗП', 'Выдано+Отгр',
			 'В ЗнО', 'Нужно еще посадить в ЗнО', 'Цель, %',
			 '% отр с ЗнО', '% отр без ЗнО']

perf_today = pd.DataFrame(data, columns = col_names, index=[today_ftd])

print('current DF created')

perf_overall = pd.read_excel('performance.xlsx',
						sep=';', decimal=',',
						encoding='ANSI',
						index_col=0)

print('overall DF uploaded')

performance = pd.concat([perf_overall, perf_today])

print('overall DF updated')

# Delete last row in the DataFrame if last two indicies(dates) are equivalent
if performance.iloc[-1].name == performance.iloc[-2].name and performance.iloc[-1].name != None :
	performance = performance[:-1]


# Ploting and save data into .png file
plt.figure(figsize=(10, 5))
sns.set_style('whitegrid')
plt.title('Динамика исполнения показателя ПРОМО')
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
print('xlsx formated')

print('done')

# with pd.ExcelWriter('performance.xlsx', mode='a') as writer:
# 	performance.to_excel(writer, sheet_name='Sheet1')
""""""