from Functions.Functions import *

MAS = min_max_stock()  # Выполняем функцию для расчета мин и макс запасов

# Обьеденение Мин и Макс запасов до уровня номенклатур
MAS = MAS.groupby('Код').agg({'ЦФО': 'first',
                              'Наименование номенклатуры': 'first',
                              'Склад хранения': 'first',
                              'Един.измерения': 'first',
                              'Поставщик': 'first',
                              'Номер договора': 'first',
                              'Кратность закупки': 'first',
                              'Минимальный заказ': 'first',
                              'Склад доставки с 1С': 'first',
                              'Склад доставки': 'first',
                              'Мин запас': 'sum',
                              'Макс запас': 'sum'}).reset_index()
MAS2 = MAS.copy()
# Готовим остатки
file = r'C:\Users\Alekseenko.v\PycharmProjects\PythonPurchaseDepartment\Расчет количества в закупку\Данные\Остатки.xlsx'  # Имя вайла с остатками
Rso_ost = stock_in_stock('Номенклатура', file)  # Выполняем функцию, которая создает DataFrame остатков сразу групируя их до уровня номенклатур
Aho_ost = stock_in_stock('Склад', file)  # Выполняем функцию, которая создает DataFrame остатков сразу групируя их до уровня складов
Aho_ost = Aho_ost[Aho_ost['Склад'] == 'ДНЕПР РЦ мелиоративное']
Aho = MAS[MAS['ЦФО'] == 'АХО']
Rso = MAS[MAS['ЦФО'] == 'НОВС/РСО']
Aho = pd.merge(Aho, Aho_ost, how='left', left_on=['Код', 'Склад хранения'], right_on=['Код', 'Склад'])
Aho.drop('Склад', inplace=True, axis=1)
Rso = pd.merge(Rso, Rso_ost, how='left', on='Код')
MAS = pd.concat([Aho, Rso], ignore_index=True)

# К закупке
MAS['В закупку'] = MAS.apply(lambda x: 0 if x['Мин запас'] <= x['Свободный остаток'] + x['Заказано у поставщиков'] else x['Макс запас'] - x['Свободный остаток'] - x['Заказано у поставщиков'], axis=1)
MAS = MAS[MAS['В закупку'] > 0]

# пересчет в закупку с учетом минимального заказа у поставщика
MAS['В закупку'] = MAS.apply(lambda x: math.ceil(x['В закупку'] / x['Минимальный заказ']) * x['Минимальный заказ'], axis=1)

# Выгрузка результатов в эксель
writer = pd.ExcelWriter('Output/Закупка от ' + str(date.today()) + '.xlsx', engine='xlsxwriter')
MAS.to_excel(writer, 'Sheet1', index=False)
Aho.to_excel(writer, 'Остаток_АХО')
Rso.to_excel(writer, 'Остаток_РСО')
MAS2.to_excel(writer, 'MAS2')
writer.save()
