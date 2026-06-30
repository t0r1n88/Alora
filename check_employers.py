"""
Скрипт для обработки списка по сотрудникам с разделением по организациям
"""

import pandas as pd
import re



df = pd.read_excel('data/Информация о сотрудниках.xlsx',dtype=str,skiprows=3,header=None)
# очищаем
df = df[df[0].notna()]
df = df[df[0] != '№ п/п']
df = df[df[0] != 'Статус организации: Функционирует']



lst_cols = ['№ п/п','ФИО','Дата рождения',
            'Пол','Гражданство','Тип документа',
            'Серия','Номер','Кем выдан',
            'Дата выдачи','Код подразделения','Место рождения',
            'ИНН','Телефон','Email',
            'Адрес регистрации','Адрес проживания','Адрес регистрации по месту пребывания',
            'Образование','Квал. Категория','Учёная степень',
            'Учёное звание ','Имеет педагогическое образование','Должность',
            'Общий стаж','Пед стаж','№ приказа о приеме на должность',
            'Дата приказа о приеме на должность','Действует с','Вид договора',
            'Ставка','Категория работника','№ приказа об увольнении',
            'Дата приказа об увольнении','Действует с','Звание',
            'Специальность','Годность','Группа учёта ',
            'Запас','№ военного билета','Состав',
            'Стоит на специальном учёте','Имеет военную подготовку','Наименование отдела ОВК',
            ]

df.columns = lst_cols

# Заполняем
df['_Организация'] = df['№ п/п'].where(~df['№ п/п'].astype(str).str.isdigit(), None)
df['_Организация'] = df['_Организация'].fillna(method='ffill')

df = df[df['№ п/п'].str.isdigit()]
df.insert(0,'Организация',df['_Организация'])
df.drop(columns=['_Организация'],inplace=True)

lst_org = df['Организация'].unique() # уникальные организации

# Делаем свод по сотрудникам
svod_df = df['Организация'].value_counts().to_frame().rename(columns={'count':'Количество сотрудников'})
svod_df = svod_df.reset_index()

# Дубликаты
dupl_df = df[df['ФИО'].duplicated(keep=False)]  # получаем дубликаты
dupl_df = dupl_df.sort_values(by='ФИО')
dupl_df.insert(0,'_ФИО',dupl_df['ФИО'])
dupl_df.drop(columns=['№ п/п','ФИО'],inplace=True)
dupl_df.rename(columns={'_ФИО':'ФИО'},inplace=True)

# Администраторы
adm_df = df[df['ФИО'].str.contains('Администратор',case=False, na=False)]
adm_df.insert(0,'_ФИО',adm_df['ФИО'])
adm_df.drop(columns=['№ п/п','ФИО'],inplace=True)
adm_df.rename(columns={'_ФИО':'ФИО'},inplace=True)





for org in lst_org:
    temp_df = df[df['Организация'] == org]
    org = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', org)
    org = org.split('(')[0].strip()
    temp_df.to_excel(f'data/Сотрудники по организациям/{org[:30]}.xlsx',index=False)
#
#
# with pd.ExcelWriter(f'data/ОБШАЯ ТАБЛИЦА сотрудников.xlsx') as writer:
#     for idx,org in enumerate(lst_org):
#         temp_df = df[df['Организация'] == org]
#         org = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', org)
#         org = org.split('(')[0].strip()
#         temp_df.to_excel(writer,sheet_name=org[:30],index=False)


with pd.ExcelWriter(f'data/Свод по сотрудникам.xlsx') as writer:
    df.to_excel(writer,sheet_name='Общая таблица',index=False)
    svod_df.to_excel(writer,sheet_name='Количество по ПОО',index=False)
    dupl_df.to_excel(writer,sheet_name='Числятся в двух и более ПОО',index=False)
    adm_df.to_excel(writer,sheet_name='Администраторы',index=False)




print('Lindy Booth')