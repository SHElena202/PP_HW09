import pandas
import pandas as pd
from matplotlib import pyplot as plt

#Для отражения всех столбцов
#pd.set_option('display.max_rows', None)
#pd.set_option('display.max_columns', None)
#pd.set_option('display.max_colwidth', None)

#Чтение таблицы
price_df = pandas.read_excel('price.xlsx')

#Вывод данных заруженной таблицы
print("\n\nДанные на дату : \n {}".format(price_df))

#Информация по столбцам:
print("\n\nИнформация по столбцам таблицы на дату: \n {}".format(price_df.info()))

#Описание данных
print("\n\nСтатистика по таблице на дату : \n {}".format(price_df.describe()))

#Выявление дубликатов
Dup_Rows = price_df[price_df.duplicated()]
print("\n\nПовторяющиеся строки : \n {}".format(Dup_Rows))

#Удаление дубликатов, сохранив только последнее вхождение
DF_RM_DUP = price_df.drop_duplicates(keep='last')
print('\n\nРезультирующий кадр данных после удаления дубликата :\n', DF_RM_DUP.head(n=5))

#Максимальная стоимость товара
print("\n\nМаксимальная стоимость товара на дату: \n {}".format(price_df[price_df['Price_euros'] == price_df['Price_euros'].max()]))

#Минимальная стоимость товара
print("\n\nМинимальная стоимость товара на дату: \n {}".format(price_df[price_df['Price_euros'] == price_df['Price_euros'].min()]))

#Малый вес, максимальные RAM и memory
print("\n\nВыборка по весу и памяти на дату: \n {}".format(
    price_df[price_df['Weight'] == price_df['Weight'].min()],
    price_df[price_df['Ram'] == price_df['Ram'].max()],
    price_df[price_df['Memory'] == price_df['Memory'].max()],
))

#Сводная таблица портативной техники
price_df1 = pd.pivot_table(price_df, index=['Company', 'Product'], values=['Price_euros', 'Weight', 'Ram'])
print("\n\nСводная портативной техники на дату  : \n {}".format(price_df1))

#Сводная таблица портативной техники OpSys
price_df2 = pd.pivot_table(price_df, index=['Company', 'Product', 'OpSys'], values=['Price_euros'])
print("\n\nСводная портативной техники OpSys на дату  : \n {}".format(price_df2))

#Графики
price_df3 = price_df.loc[((price_df['Company'] == 'HP') & (price_df['TypeName'] == 'Notebook'))]
price_df4 = pd.pivot_table(price_df3, index=['Product'], columns=['TypeName'], values=['Price_euros'])
print("\n\nСводная HP Notebook на дату : \n {}".format(price_df4))


#Линейный график по умолчанию
price_df4.plot()
plt.xticks(rotation=45)
plt.show()

#Гистограмма
price_df4.plot(kind='bar')
plt.xticks(rotation=90)
plt.show()

#Сбор данных в новый файл
price_sheets = {'Price': price_df1}
writer1 = pd.ExcelWriter(r'C:\Users\selivanova\Desktop\HM Python\HW09\Price1.xlsx', engine='xlsxwriter')
for sheet_name in price_sheets.keys():
    price_sheets[sheet_name].to_excel(writer1, sheet_name=sheet_name, index=True)

writer1.save()
