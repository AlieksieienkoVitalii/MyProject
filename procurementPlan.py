from Functions import *

df_mas = min_max_stock()  # Выполняем функцию для расчета мин и макс запасов
file = r'C:\Users\Alekseenko.v\PycharmProjects\PythonPurchaseDepartment\Расчет количества в закупку\Данные\Остатки.xlsx'  # Имя файла с остатками
df_stock = stock_in_stock('Склад', file)  # Выполняем функцию, которая создает DataFrame остатков сразу групируя их до уровня складов
df_mas = pd.merge(df_mas, df_stock, how='left', left_on=['Код', 'Склад хранения'], right_on=['Код', 'Склад'])  # Подтягиваем остатки к DataFrame с мин и мас запасами

df_mas.fillna({'Свободный остаток': 0, 'Заказано у поставщиков': 0}, inplace=True)  # Заменяем пустые значения на "0"
df_surplus = df_mas.copy()  # Создаем новый DataFarme

# Расчет потребности
df_mas['Потребность'] = df_mas.apply(lambda x: 0 if x['Мин запас'] <= x['Свободный остаток'] else x['Макс запас'] - x['Свободный остаток'], axis=1)
Out_of_stock = df_mas.copy()
df_mas = df_mas[df_mas['Потребность'] > 0]

# Пересчет потребности с учетом кратности
df_mas['Потребность'] = df_mas.apply(lambda x: math.ceil(x['Потребность'] / x['Кратность закупки']) * x['Кратность закупки'], axis=1)


# Подготовка DataFrame с излишками
df_surplus['Излишки'] = df_surplus.apply(lambda x: x['Свободный остаток'] - x['Макс запас'] if x['Свободный остаток'] > x['Макс запас'] else 0, axis=1)  # Расчет излишков
df_surplus = df_surplus[df_surplus['Излишки'] != 0]  # Убираем строки, по которым нет излишков
df_surplus = df_surplus[df_surplus['ЦФО'] != 'АХО']  # Убираем строки, которые относятся к ЦФО АХО
df_surplus = df_surplus[['Код', 'Склад хранения', 'Кратность закупки', 'Излишки']]  # Удаляем лишние столбцы и меняем последовательность
df_surplus = df_surplus[::].reset_index(drop=True)  # Делаем нумирацию индексов последовательной

# Пересчет излишков с учетом кратности
df_surplus['Излишки'] = df_surplus.apply(lambda x: int(x['Излишки'] / x['Кратность закупки']) * x['Кратность закупки'], axis=1)

# Перереаспределение остатков
Result = pd.DataFrame({'Дата доставки': [],
                       'Склад Отправитель': [],
                       'Склад Получатель': [],
                       'Код': [],
                       'Качество': [],
                       'Кол-во': [],
                       'Комментарий': [],
                       'Номер Битрикс': []})  # Создаем пустую DataFrame для наполнения перемещениями
df_mas['Количество'] = 0
df_mas['Склад отправитель'] = ''
df_mas['В закупку'] = df_mas['Потребность']

df_mas = df_mas[::].reset_index(drop=True)  # Делаем нумирацию индексов последовательной

for i in range(len(df_mas)):
    check = df_surplus.index[df_surplus['Код'] == df_mas.loc[i, 'Код']].tolist()  # Получаем лист с перечнем индексов с файла излишков которые соответствуют условию поиска
    if check:
        for ii in range(len(check)):
            key = df_mas.loc[i, 'Код']
            sklad_pol = df_mas.loc[i, 'Склад хранения']
            sklad_otp = df_surplus.loc[check[ii], 'Склад хранения']
            if df_mas.loc[i, 'В закупку'] >= df_surplus.loc[check[ii], 'Излишки']:
                kolvo = df_surplus.loc[check[ii], 'Излишки']
                df_surplus.loc[check[ii], 'Излишки'] = 0
                df_surplus.loc[check[ii], 'Код'] = 'Пусто'
                df_mas.loc[i, 'В закупку'] = df_mas.loc[i, 'В закупку'] - kolvo
                df_mas.loc[i, 'Количество'] = df_mas.loc[i, 'Количество'] + kolvo
                df_mas.loc[i, 'Склад отправитель'] = df_mas.loc[i, 'Склад отправитель'] + '___' + sklad_otp + ' - ' + str(kolvo)
            elif df_mas.loc[i, 'В закупку'] < df_surplus.loc[check[ii], 'Излишки']:
                 kolvo = df_mas.loc[i, 'В закупку']
                 df_mas.loc[i, 'В закупку'] = 0
                 df_mas.loc[i, 'Количество'] = df_mas.loc[i, 'Количество'] + kolvo
                 df_mas.loc[i, 'Склад отправитель'] = df_mas.loc[i, 'Склад отправитель'] + '___' + sklad_otp + ' - ' + str(kolvo)
                 df_surplus.loc[check[ii], 'Излишки'] = df_surplus.loc[check[ii], 'Излишки'] - kolvo
                 Result = Result.append([{'Дата доставки': '',
                                          'Склад Отправитель': sklad_otp,
                                          'Склад Получатель': sklad_pol,
                                          'Код': key,
                                          'Качество': '',
                                          'Кол-во': kolvo,
                                          'Комментарий': '',
                                          'Номер Битрикс': ''}], ignore_index=True)
                 break
            else:
                print('Непредвиденое условие при перераспределении остатков, посмотри в код!!!')
            Result = Result.append([{'Дата доставки': '',
                                    'Склад Отправитель': sklad_otp,
                                    'Склад Получатель': sklad_pol,
                                    'Код': key,
                                    'Качество': '',
                                    'Кол-во': kolvo,
                                    'Комментарий': '',
                                    'Номер Битрикс': ''}], ignore_index=True)

# Готовим файл к закупке
zakypka = df_mas[['ЦФО',
                  'Код',
                  'Наименование номенклатуры',
                  'Един.измерения',
                  'Минимальный заказ',
                  'В закупку']]  # Удаляем лишние столбцы и меняем последовательность
zakypka = zakypka[zakypka['В закупку'] != 0]

# Группируем до уровня номенклатур
zakypka = zakypka.groupby('Код').agg({'Наименование номенклатуры': 'first',
                                      'Един.измерения': 'first',
                                      'Минимальный заказ': 'first',
                                      'ЦФО': 'first',
                                      'В закупку': 'sum'}).reset_index()

df_stock_nomenklatura = stock_in_stock('Номенклатура', file)  # Выполняем функцию, которая создает DataFrame остатков сразу групируя их до уровня складов
zakypka = pd.merge(zakypka, df_stock_nomenklatura, how='left', left_on=['Код'], right_on=['Код'])
zakypka['В закупку'] = zakypka.apply(lambda x: 0 if x['В закупку'] <= x['Заказано у поставщиков'] else x['В закупку'] - x['Заказано у поставщиков'], axis=1)
 # пересчет Потребность с учетом минимального заказа у поставщика
zakypka['В закупку'] = zakypka.apply(lambda x: x['Минимальный заказ'] if x['В закупку'] < x['Минимальный заказ'] else x['В закупку'], axis=1)
zakypka = zakypka[['ЦФО',
                  'Код',
                  'Наименование номенклатуры',
                  'Един.измерения',
                  'В закупку']]  # Удаляем лишние столбцы и меняем последовательность


# Выгрузка результатов в эксель
writer = pd.ExcelWriter('План закупок от ' + str(date.today()) + '.xlsx', engine='xlsxwriter')
# df_mas.to_excel(writer, 'Потребность')
# df_surplus.to_excel(writer, 'Излишки')
zakypka.to_excel(writer, 'Закупка', index=False)
Result.to_excel(writer, 'Перемещения', index=False)
Out_of_stock.to_excel(writer, 'Out_of_stock', index=False)
writer.save()



# writer = pd.ExcelWriter('План.xlsx', engine='xlsxwriter')
# df_surplus.to_excel(writer, 'Излишки')
# writer.save()
