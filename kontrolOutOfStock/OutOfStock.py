from Functions import *
from reallocationBetweenStores import *

df_mas = min_max_stock()  # Выполняем функцию для расчета мин и макс запасов
file = r'C:\Users\Alekseenko.v\PycharmProjects\PythonPurchaseDepartment\Расчет количества в закупку\Данные\Остатки.xlsx'  # Имя вайла с остатками
df_stock = stock_in_stock('Склад', file)  # Выполняем функцию, которая создает DataFrame остатков сразу групируя их до уровня складов
df_mas = pd.merge(df_mas, df_stock, how='left', left_on=['Код', 'Склад хранения'], right_on=['Код', 'Склад'])  # Подтягиваем остатки к DataFrame с мин и мас запасами
df_mas.fillna({'Свободный остаток': 0, 'Заказано у поставщиков': 0}, inplace=True)  # Заменяем пустые значения на "0"
# df_mas = df_mas[df_mas['Свободный остаток'] == 0]
df_mas['Дата закупки'] = df_mas.apply(lambda x: date.today() + timedelta(x['Плече поставки']), axis=1)

df_surplus = reallocation_between_stores()


# Выгрузка результатов в эксель
writer = pd.ExcelWriter('Out-of-stock от ' + str(date.today()) + '.xlsx', engine='xlsxwriter')
df_mas.to_excel(writer, 'Out-of-stock')
df_surplus.to_excel(writer, 'df_surplus')
writer.save()