import os
import pandas as pd

path = r'C:\Users\Alekseenko.v\Desktop\Инфо_по_складам'
files = os.listdir(path)
df = pd.DataFrame()
for f in files:
    data = pd.read_excel(r'C:\Users\Alekseenko.v\Desktop\Инфо_по_складам\\' + f, 'Sheet1')
    data['Склад'] = f[:-4]
    df = df.append(data)
df.drop(df.columns[[5, 8, 11]], axis=1, inplace=True)
df[['К-во дней с посл.расхода']] = df[['К-во дней с посл.расхода']].fillna(value=0)
df = df[::].reset_index(drop=True)  # Делаем нумирацию индексов последовательной


def years_months_days(totaldays):
    years = totaldays // 365.25
    months = (totaldays % 365.25) // 30.4375
    days = (totaldays % 365.25) % 30.4375
    return [years, months, days]


# Определение периода без движения
for i in range(len(df)):
    yearsMonthsDays = years_months_days(df.loc[i, 'К-во дней с посл.расхода'])
    if yearsMonthsDays[0] == 0 and yearsMonthsDays[1] == 0 and yearsMonthsDays[2] == 0:
        df.loc[i, 'Движение'] = 'Движение отсутствует последние - ' + df.columns[5][20:]
    elif yearsMonthsDays[0] > 0:
        period = 'год'
        if 1 < yearsMonthsDays[0] <= 4:
            period = 'года'
        elif yearsMonthsDays[0] > 4:
            period = 'лет'
        df.loc[i, 'Движение'] = 'Движение отсутствует ' + str(int(yearsMonthsDays[0])) + ' ' + period
    elif yearsMonthsDays[0] == 0 and 12 >= yearsMonthsDays[1] > 0:
        period = 'месяц'
        if 1 < yearsMonthsDays[1] <= 4:
            period = 'месяца'
        elif 4 < yearsMonthsDays[1] <= 12:
            period = 'месяцев'
        df.loc[i, 'Движение'] = 'Движение отсутствует ' + str(int(yearsMonthsDays[1])) + ' ' + period
    elif yearsMonthsDays[0] == 0 and yearsMonthsDays[1] == 0 and 30.4375 >= yearsMonthsDays[2] > 0:
        df.loc[i, 'Движение'] = 'Движение присутствует за последние 30 дней'

# Расчет прогнозного периода запаса
for i in range(len(df)):
    if df.loc[i, 'Запас ТМЦ, дни'] == ' -':
        df.loc[i, 'Запас'] = 'data отсутствуют'
        continue
    yearsMonthsDays = years_months_days(int(df.loc[i, 'Запас ТМЦ, дни']))
    if yearsMonthsDays[0] == 0 and yearsMonthsDays[1] == 0 and yearsMonthsDays[2] == 0:
        df.loc[i, 'Запас'] = 'Запас отсутствует'
    elif yearsMonthsDays[0] > 0:
        period = 'год'
        if 1 < yearsMonthsDays[0] <= 4:
            period = 'года'
        elif yearsMonthsDays[0] > 4:
            period = 'лет'
        if yearsMonthsDays[0] > 5:
            df.loc[i, 'Запас'] = 'Запас более 5 лет'
        else:
            df.loc[i, 'Запас'] = 'Запас на - ' + str(int(yearsMonthsDays[0])) + ' ' + period
    elif yearsMonthsDays[0] == 0 and 12 >= yearsMonthsDays[1] > 0:
        period = 'месяц'
        if 1 < yearsMonthsDays[1] <= 4:
            period = 'месяца'
        elif 4 < yearsMonthsDays[1] <= 12:
            period = 'месяцев'
        df.loc[i, 'Запас'] = 'Запас на - ' + str(int(yearsMonthsDays[1])) + ' ' + period
    elif yearsMonthsDays[0] == 0 and yearsMonthsDays[1] == 0 and 30.4375 >= yearsMonthsDays[2] > 0:
        df.loc[i, 'Запас'] = 'Запас на месяц'


# Выгрузка результатов в эксель
writer = pd.ExcelWriter('Оборачиваемость.xlsx', engine='xlsxwriter')
df.to_excel(writer, 'Лист1', index=False)
writer.save()