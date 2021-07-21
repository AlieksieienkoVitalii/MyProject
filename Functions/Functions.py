import pandas as pd
from datetime import date
import math
from datetime import timedelta
from workalendar.europe import Ukraine


def min_max_stock():
    name_directory = r'C:\Users\Alekseenko.v\PycharmProjects\PythonPurchaseDepartment\Расчет количества в закупку\Данные\Справочники.xlsx'

    df_b1 = pd.read_excel(name_directory, sheet_name='Блок_1', usecols='D:P', skiprows=2)
    df_b1 = df_b1[["Код", "Поставщик", "Номер договора", "Кратность закупки", "Минимальный заказ", "Принято в работу?"]]

    df_b2 = pd.read_excel(name_directory, sheet_name='Блок_2', usecols='B:K', skiprows=2)
    df_b3 = pd.read_excel(name_directory, sheet_name='Блок_3', usecols='A:J', skiprows=2)
    df_b4 = pd.read_excel(name_directory, sheet_name='Блок_4', usecols='A:B', skiprows=2)
    df_b7 = pd.read_excel(name_directory, sheet_name='Блок_7', usecols='A:B')

    # Склеивание 3х массивов в один для расчета Мин и Макс запасов
    df_mas = df_b3.merge(df_b1, how='left', left_on='Код', right_on='Код')
    df_mas = df_mas.merge(df_b2)
    df_mas = df_mas.merge(df_b4)
    df_mas.drop(['Способ получения актуального плана операций в периоде', 'Наименование поставщика в 1С'], inplace=True, axis=1)

    df_mas = df_mas[df_mas['Принято в работу?'] == True]  # Удаление строк, номенклатуры которых не приняты в работу
    df_mas['Период запаса'] = df_mas['Период запаса'].astype(float)  # Изменение типа данных с int на float
    df_mas['Код'] = df_mas['Код'].astype(float)  # Изменение типа данных с int на float

    # Календарь праздников и выходных для Украины
    cal = Ukraine()
    cal.holidays(date.today().year)

    # Расчет плеча поставки
    df_mas['Обработка менеджером'] = df_mas.apply(lambda x: float(df_b7.iloc[1, 1]) + 1 if x['Условия оплаты'] != 'Предоплата' else 0, axis=1)
    df_mas['Предоплата'] = df_mas['Условия оплаты'].apply(lambda x: float(df_b7.iloc[2, 1]) if x == 'Предоплата' else 0)
    df_mas['Срок поставки кал'] = df_mas.apply(lambda x: 0 if x['Календарные/рабочие дни'] == 'Рабочих' else x['Срок поставки (дней)'],axis=1)
    df_mas['Срок поставки раб'] = df_mas.apply(lambda x: x['Срок поставки (дней)'] if x['Календарные/рабочие дни'] == 'Рабочих' else 0, axis=1)
    df_mas['Самовывоз'] = df_mas['Склад доставки'].apply(lambda x: float(df_b7.iloc[9, 1]) if x == 'Самовывоз' else 0)
    df_mas['Перемещение'] = df_mas.apply(lambda x: float(df_b7.iloc[9, 1]) if x['Склад хранения'] != x['Склад доставки с 1С'] else 0, axis=1)
    df_mas['Плече поставки'] = df_mas.apply(lambda x: cal.add_working_days(date.today(), x['Обработка менеджером']) + timedelta(x['Предоплата']) + timedelta(x['Срок поставки кал']), axis=1)
    df_mas['Плече поставки'] = df_mas.apply(lambda x: cal.add_working_days(x['Плече поставки'], x['Срок поставки раб']) + timedelta(x['Самовывоз']) + timedelta(x['Перемещение']), axis=1)
    df_mas['Плече поставки'] = df_mas.apply(lambda x: x['Плече поставки'] - date.today(), axis=1)
    df_mas['Плече поставки'] = df_mas['Плече поставки'] / pd.to_timedelta('1D')  # Перевод даты в количество дней

    # Расчет Мин и Макс запасов
    df_mas['Коэффициент не точности планирования'] = df_mas['Коэффициент не точности планирования'].apply(lambda x: 1 if x == 0 else x)  # Замена нулей на 1
    df_mas['Мин запас'] = ((df_mas['Плановое количество драйверов в периоде запаса'] * df_mas['Норма потребления на одну операцию']) / df_mas['Период запаса'] * df_mas['Плече поставки']) * df_mas['Коэффициент не точности планирования']
    df_mas['Макс запас'] = df_mas.apply(lambda x: x['Мин запас'] / x['Плече поставки'] * (x['Плече поставки'] + 7), axis=1)

    # Обьеденение Мин и Макс запасов до уровня складов
    df_mas = df_mas.groupby(['Код', 'Склад хранения']).agg({'ЦФО': 'first',
                                                            'Наименование номенклатуры': 'first',
                                                            'Един.измерения': 'first',
                                                            'Поставщик': 'first',
                                                            'Номер договора': 'first',
                                                            'Кратность закупки': 'first',
                                                            'Минимальный заказ': 'first',
                                                            'Склад доставки с 1С': 'first',
                                                            'Склад доставки': 'first',
                                                            'Плече поставки': 'first',
                                                            'Мин запас': 'sum',
                                                            'Макс запас': 'sum'}).reset_index()
    # Пересчет потребности с учетом кратности
    # df_mas['Мин запас'] = df_mas.apply(lambda x: math.ceil(x['Мин запас'] / x['Кратность закупки']) * x['Кратность закупки'], axis=1)
    # df_mas['Макс запас'] = df_mas.apply(lambda x: math.ceil(x['Макс запас'] / x['Кратность закупки']) * x['Кратность закупки'], axis=1)
    return df_mas


def stock_in_stock(result_type, name_stock):
    df_stock = pd.read_excel(name_stock, sheet_name='TDSheet', usecols='B:K', skiprows=8)
    df_stock = df_stock.rename(columns={df_stock.columns[8]: 'Заказано у поставщиков', df_stock.columns[9]: 'Свободный остаток'})

    if result_type == 'Склад':
        df_stock = df_stock.groupby(['Код', 'Склад']).agg({'Заказано у поставщиков': 'sum', 'Свободный остаток': 'sum'}).reset_index()
    elif result_type == 'Номенклатура':
        df_stock = df_stock.groupby('Код').agg({'Заказано у поставщиков': 'sum', 'Свободный остаток': 'sum'}).reset_index()
    else:
        print('Не верно выбрат тип отчета остатков! Необходимо указать один из вариантов! "Склад"; "Номенклатура"')
    return df_stock


def surplus():  # Расчет излишков
    df_mas = min_max_stock()
    df_stock = stock_in_stock('Склад')
    df_mas = pd.merge(df_mas, df_stock, how='left', left_on=['Код', 'Склад хранения'], right_on=['Код', 'Склад'])
    df_mas['Излишки'] = df_mas.apply(
        lambda x: x['Свободный остаток'] - x['Макс запас'] if x['Свободный остаток'] > x['Макс запас'] else 0, axis=1)
    df_mas = df_mas[df_mas['Излишки'] != 0]
    df_mas = df_mas[df_mas['ЦФО'] != 'АХО']
    df_mas = df_mas[['Код', 'Склад хранения', 'Излишки']]
    df_mas.index = df_mas.apply(lambda x: str(x['Код']) + ';' + x['Склад хранения'], axis=1)
    return df_mas,




# import pandas as pd
# from datetime import date
# import math
# from datetime import timedelta
# from workalendar.europe import Ukraine
#
# name_directory = "Справочники.xlsx"
# name_stock = "Остатки.xlsx"
#
# df_b1 = pd.read_excel(name_directory, sheet_name='Блок_1', usecols='D:N', skiprows=2)
# df_b1 = df_b1[["Код", "Поставщик", "Номер договора", "Кратность закупки", "Минимальный заказ", "Принято в работу?"]]
#
# df_b2 = pd.read_excel(name_directory, sheet_name='Блок_2', usecols='B:K', skiprows=2)
# df_b3 = pd.read_excel(name_directory, sheet_name='Блок_3', usecols='A:J', skiprows=2)
# df_b4 = pd.read_excel(name_directory, sheet_name='Блок_4', usecols='A:B', skiprows=2)
# df_b7 = pd.read_excel(name_directory, sheet_name='Блок_7', usecols='A:B')
# df_stock = pd.read_excel(name_stock, sheet_name='TDSheet', usecols='B:K', skiprows=8)
#
# # Склеивание 3х массивов в один для расчета Мин и Макс запасов
# MAS = df_b3.merge(df_b1, how='left', left_on='Код', right_on='Код')
# MAS = MAS.merge(df_b2)
# MAS = MAS.merge(df_b4)
# MAS.drop(['Способ получения актуального плана операций в периоде', 'Наименование поставщика в 1С'], inplace=True, axis=1)
#
# MAS = MAS[MAS['Принято в работу?'] == True]  # Удаление строк, номенклатуры которых не приняты в работу
# MAS['Период запаса'] = MAS['Период запаса'].astype(float)  # Изменение типа данных с int на float
# MAS['Код'] = MAS['Код'].astype(float)  # Изменение типа данных с int на float
#
# # Календарь праздников и выходных для Украины
# cal = Ukraine()
# cal.holidays(date.today().year)
#
# # Расчет плеча поставки
# MAS['Обработка менеджером'] = MAS.apply(lambda x: float(df_b7.iloc[1, 1]) + 1 if x['Условия оплаты'] != 'Предоплата' else 0, axis=1)
# MAS['Предоплата'] = MAS['Условия оплаты'].apply(lambda x: float(df_b7.iloc[2, 1]) if x == 'Предоплата' else 0)
# MAS['Срок поставки кал'] = MAS.apply(lambda x: 0 if x['Календарные/рабочие дни'] == 'Рабочих' else x['Срок поставки (дней)'],axis=1)
# MAS['Срок поставки раб'] = MAS.apply(lambda x: x['Срок поставки (дней)'] if x['Календарные/рабочие дни'] == 'Рабочих' else 0, axis=1)
# MAS['Самовывоз'] = MAS['Склад доставки'].apply(lambda x: float(df_b7.iloc[9, 1]) if x == 'Самовывоз' else 0)
# MAS['Перемещение'] = MAS.apply(lambda x: float(df_b7.iloc[9, 1]) if x['Склад хранения'] != x['Склад доставки с 1С'] else 0, axis=1)
# MAS['Плече поставки'] = MAS.apply(lambda x: cal.add_working_days(date.today(), x['Обработка менеджером']) + timedelta(x['Предоплата']) + timedelta(x['Срок поставки кал']), axis=1)
# MAS['Плече поставки'] = MAS.apply(lambda x: cal.add_working_days(x['Плече поставки'], x['Срок поставки раб']) + timedelta(x['Самовывоз']) + timedelta(x['Перемещение']), axis=1)
# MAS['Плече поставки'] = MAS.apply(lambda x: x['Плече поставки'] - date.today(), axis=1)
# MAS['Плече поставки'] = MAS['Плече поставки'] / pd.to_timedelta('1D')  # Перевод даты в количество дней
#
# # Расчет Мин и Макс запасов
# MAS['Коэффициент не точности планирования'] = MAS['Коэффициент не точности планирования'].apply(lambda x: 1 if x == 0 else x)  # Замена нулей на 1
# MAS['Мин запас'] = ((MAS['Плановое количество драйверов в периоде запаса'] * MAS['Норма потребления на одну операцию']) / MAS['Период запаса'] * MAS['Плече поставки']) * MAS['Коэффициент не точности планирования']
# MAS['Макс запас'] = MAS.apply(lambda x: x['Мин запас'] / x['Плече поставки'] * (x['Плече поставки'] + 7), axis=1)
#
# MAS2 = MAS.copy()  # Удалить !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#
# # Обьеденение Мин и Макс запасов до уровня складов
# MAS = MAS.groupby(['Код', 'Склад хранения']).agg({'ЦФО': 'first',
#                                                   'Наименование номенклатуры': 'first',
#                                                   'Един.измерения': 'first',
#                                                   'Поставщик': 'first',
#                                                   'Номер договора': 'first',
#                                                   'Кратность закупки': 'first',
#                                                   'Минимальный заказ': 'first',
#                                                   'Склад доставки с 1С': 'first',
#                                                   'Склад доставки': 'first',
#                                                   'Мин запас': 'sum',
#                                                   'Макс запас': 'sum'}).reset_index()
# # Пересчет потребности с учетом кратности
# MAS['Мин запас'] = MAS.apply(lambda x: math.ceil(x['Мин запас'] / x['Кратность закупки']) * x['Кратность закупки'], axis=1)
# MAS['Макс запас'] = MAS.apply(lambda x: math.ceil(x['Макс запас'] / x['Кратность закупки']) * x['Кратность закупки'], axis=1)
#
#
# # Обьеденение Мин и Макс запасов до уровня номенклатур
# MAS = MAS.groupby('Код').agg({'ЦФО': 'first',
#                               'Наименование номенклатуры': 'first',
#                               'Склад хранения': 'first',
#                               'Един.измерения': 'first',
#                               'Поставщик': 'first',
#                               'Номер договора': 'first',
#                               'Кратность закупки': 'first',
#                               'Минимальный заказ': 'first',
#                               'Склад доставки с 1С': 'first',
#                               'Склад доставки': 'first',
#                               'Мин запас': 'sum',
#                               'Макс запас': 'sum'}).reset_index()
#
# # Готовим остатки
# df_stock = df_stock.rename(columns={df_stock.columns[8]: 'Заказано у поставщиков', df_stock.columns[9]: 'Свободный остаток'})
# Rso_ost = df_stock.groupby('Код').agg({'Заказано у поставщиков': 'sum',
#                                        'Свободный остаток': 'sum'}).reset_index()
# Aho_ost = df_stock.groupby(['Код', 'Склад']).agg({'Заказано у поставщиков': 'sum',
#                                                   'Свободный остаток': 'sum'}).reset_index()
# Aho_ost = Aho_ost[Aho_ost['Склад'] == 'ДНЕПР РЦ мелиоративное']
# Aho = MAS[MAS['ЦФО'] == 'АХО']
# Rso = MAS[MAS['ЦФО'] == 'НОВС/РСО']
# Aho = pd.merge(Aho, Aho_ost, how='left', left_on=['Код', 'Склад хранения'], right_on=['Код', 'Склад'])
# Aho.drop('Склад', inplace=True, axis=1)
# Rso = pd.merge(Rso, Rso_ost, how='left', on='Код')
# MAS = pd.concat([Aho, Rso], ignore_index=True)
#
# # К закупке
# MAS['В закупку'] = MAS.apply(lambda x: 0 if x['Мин запас'] <= x['Свободный остаток'] + x['Заказано у поставщиков'] else x['Макс запас'] - x['Свободный остаток'] - x['Заказано у поставщиков'], axis=1)
# MAS = MAS[MAS['В закупку'] > 0]
#
# # пересчет в закупку с учетом минимального заказа у поставщика
# MAS['В закупку'] = MAS.apply(lambda x: math.ceil(x['В закупку'] / x['Минимальный заказ']) * x['Минимальный заказ'], axis=1)
#
#
# # Выгрузка результатов в эксель
# writer = pd.ExcelWriter('Закупка от ' + str(date.today()) + '.xlsx', engine='xlsxwriter')
# MAS.to_excel(writer, 'Sheet1', index=False)
# Aho.to_excel(writer, 'Остаток_АХО')
# Rso.to_excel(writer, 'Остаток_РСО')
# MAS2.to_excel(writer, 'Удалить')  # Удалить !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# writer.save()
