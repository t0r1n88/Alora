"""
Скрипт для проверки выгрузки гражданства
"""

import pandas as pd
import openpyxl

def diff_age_zach(row:pd.Series):
    """
    Функция для вычисления ошибки когда между датой рождения и датой зачисления меньше 14 лет
    :param row:
    :return:
    """
    date_zach, date_birth = row

    result = date_zach.year - date_birth.year

    if result < 14:
        return 'Ошибка: Разница между датой рождения и датой зачисления меньше 14 лет'
    else:
        return None


def calc_end(value):
    if value.year < 2026:
        return 'Ошибка: Дата завершения обучения ранее 2026'
    else:
        return None

def diff_otch_zach(row:pd.Series):
    date_zach, date_otch = row

    if date_zach >= date_otch:
        return 'Ошибка: Дата зачисления равна или больше чем дата завершения обучения'
    else:
        return None




df = pd.read_excel('data/Контроль по студентам.xlsx',skiprows=3,header=None,dtype=str)

# очищаем
df = df[df[0].notna()]
df = df[df[0] != '№ п/п']

df.columns = ['N п/п','ФИО','Дата рождения','Возраст','Гражданство','Тип документа удостоверяющего личность',
              'Серия документа удостоверяющего личность','Номер документа удостоверяющего личность',
              'Код специальности','Название специальности','Дата зачисления','Предполагаемая дата завершения обучения']

# Заполняем
df['_Организация'] = df['N п/п'].where(~df['N п/п'].astype(str).str.isdigit(), None)
df['_Организация'] = df['_Организация'].fillna(method='ffill')

df = df[df['N п/п'].str.isdigit()]
df.insert(0,'Организация',df['_Организация'])
df.drop(columns=['_Организация'],inplace=True)

# Подготавливаем колонки
for col in ['Дата рождения', 'Дата зачисления', 'Предполагаемая дата завершения обучения']:
    df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
df['Возраст'] = df['Возраст'].astype(int)

# Ищем ошибки
df['Ошибка возраст'] = df['Возраст'].apply(lambda x:'Ошибка: Меньше 15 лет' if x < 15 else None)
df['Ошибка Разница Рождение Зачисление < 14'] = df[['Дата зачисления','Дата рождения']].apply(diff_age_zach,axis=1)

df['Ошибка Дата завершения не 2026'] = df['Предполагаемая дата завершения обучения'].apply(calc_end)

df['Ошибка Зачисление Отчисление'] =  df[['Дата зачисления','Предполагаемая дата завершения обучения']].apply(diff_otch_zach,axis=1)






df.to_excel('data/res.xlsx',index=False)




print('Lindy Booth')





