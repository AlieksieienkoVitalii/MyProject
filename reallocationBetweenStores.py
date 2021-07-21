from Functions import *


def reallocation_between_stores():

    df_mas = min_max_stock()  # Выполняем функцию для расчета мин и макс запасов
    file = r'C:\Users\Alekseenko.v\PycharmProjects\PythonPurchaseDepartment\Расчет количества в закупку\Данные\Остатки.xlsx'  # Имя файла с остатками
    df_stock = stock_in_stock('Склад', file)  # Выполняем функцию, которая создает DataFrame остатков сразу групируя их до уровня складов
    df_mas = pd.merge(df_mas, df_stock, how='left', left_on=['Код', 'Склад хранения'], right_on=['Код', 'Склад'])  # Подтягиваем остатки к DataFrame с мин и мас запасами
    df_mas.fillna({'Свободный остаток': 0, 'Заказано у поставщиков': 0}, inplace=True)  # Заменяем пустые значения на "0"
    df_surplus = df_mas.copy()  # Создаем новый DataFarme

    # Подготовка DataFrame с излишками
    df_surplus['Излишки'] = df_surplus.apply(lambda x: x['Свободный остаток'] - x['Макс запас'] if x['Свободный остаток'] > x['Макс запас'] else 0, axis=1)  # Расчет излишков
    df_surplus = df_surplus[df_surplus['Излишки'] != 0]  # Убираем строки, по которым нет излишков
    df_surplus = df_surplus[df_surplus['ЦФО'] != 'АХО']  # Убираем строки, которые относятся к ЦФО АХО
    df_surplus = df_surplus[['Код', 'Склад хранения', 'Излишки']]  # Удаляем лишние столбцы и меняем последовательность
    df_surplus = df_surplus[::].reset_index(drop=True)  # Делаем нумирацию индексов последовательной

    # Подготовка DataFrame с потребностью
    df_mas = df_mas[['ЦФО',
                     'Склад хранения',
                     'Код',
                     'Наименование номенклатуры',
                     'Един.измерения',
                     'Мин запас',
                     'Макс запас',
                     'Заказано у поставщиков',
                     'Свободный остаток']]  # Удаляем лишние столбцы и меняем последовательность
    df_mas['Потребность'] = df_mas.apply(lambda x: x['Макс запас'] - x['Свободный остаток'] - x['Заказано у поставщиков'] if x['Мин запас'] > x['Свободный остаток'] + x['Заказано у поставщиков'] else 0, axis=1)
    df_mas = df_mas[df_mas['Потребность'] != 0]
    df_mas = df_mas[::].reset_index(drop=True)  # Делаем нумирацию индексов последовательной

    otchet_df_mas = df_mas.copy()
    otchet_df_surplus = df_surplus.copy()

    # Перереаспределение остатков
    Result = pd.DataFrame({'Дата доставки': [],
                           'Склад Отправитель': [],
                           'Склад Получатель': [],
                           'Код': [],
                           'Качество': [],
                           'Кол-во': [],
                           'Комментарий': [],
                           'Номер Битрикс': []})  # Создаем пустую DataFrame для наполнения перемещениями
    for i in range(len(df_mas)):
        check = df_surplus.index[df_surplus['Код'] == df_mas.loc[i, 'Код']].tolist()  # Получаем лист с перечнем индексов с файла излишков которые соответствуют условию поиска
        if check:
            for ii in range(len(check)):
                key = df_mas.loc[i, 'Код']
                sklad_pol = df_mas.loc[i, 'Склад хранения']
                sklad_otp = df_surplus.loc[check[ii], 'Склад хранения']
                if df_mas.loc[i, 'Потребность'] >= df_surplus.loc[check[ii], 'Излишки']:
                    kolvo = df_surplus.loc[check[ii], 'Излишки']
                    df_surplus.loc[check[ii], 'Излишки'] = 0
                    df_surplus.loc[check[ii], 'Код'] = 'Пусто'
                    df_mas.loc[i, 'Потребность'] = df_mas.loc[i, 'Потребность'] - kolvo
                elif df_mas.loc[i, 'Потребность'] < df_surplus.loc[check[ii], 'Излишки']:
                    kolvo = df_mas.loc[i, 'Потребность']
                    df_mas.loc[i, 'Потребность'] = 0
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
    return Result


Result = reallocation_between_stores()
# Выгрузка результатов в эксель
writer = pd.ExcelWriter('PereraspredeleniyeOstatkov от ' + str(date.today()) + '.xlsx', engine='xlsxwriter')
Result.to_excel(writer, 'Автозаливка', index=False)

writer.save()

