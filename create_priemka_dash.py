"""
Скрипт для подготовки данных к загрузке в дашборд по приемной кампании
"""
import re

import pandas as pd
import time



def extract_educ_form(value):
    if pd.isna(value):
        return 'Не заполнена форма обучения'

    form_str = str(value).strip()

    if re.search(r'\bочная\b',form_str,re.IGNORECASE):
        return 'очная'
    elif re.search(r'\очно[- ]заочн|очн[оы]е?[- ]заочн|оч[-.]заоч',form_str,re.IGNORECASE):
        return 'очно-заочная'
    elif re.search(r'\bзаочная\b',form_str,re.IGNORECASE):
        return 'заочная'

    else:
        return f'{value} неизвестная форма обучения'


def extract_pay_form(value):
    if pd.isna(value):
        return 'Не заполнена форма оплаты'

    form_str = str(value).strip()

    if re.search(r'внебюджет',form_str,re.IGNORECASE):
        return 'коммерческая'
    elif re.search(r'бюджет',form_str,re.IGNORECASE):
        return 'бюджет'
    elif re.search(r'коммер|платн|оплат|договор',form_str,re.IGNORECASE):
        return 'коммерческая'

    else:
        return f'{value} неизвестная форма оплаты'


def extract_level_educ(value):
    if pd.isna(value):
        return 'Не заполнен уровень образования'

    form_str = str(value).strip()

    if re.search(r'среднее профессиональное образование',form_str,re.IGNORECASE):
        return 'среднее профессиональное образование'
    elif re.search(r'основное',form_str,re.IGNORECASE):
        return 'основное общее'
    elif re.search(r'среднее',form_str,re.IGNORECASE):
        return 'среднее общее'
    elif re.search(r'специальное коррекционное образование',form_str,re.IGNORECASE):
        return 'специальное коррекционное образование'

    else:
        return f'{value} неизвестный уровень образования'





def generate_data_for_dash_priem(data_file:str,end_folder:str):
    """

    :param data_file: файл выгрузки
    :param end_folder: конечная папка
    :return:
    """
    # Лист орг-я и спец-ть СПО
    org_spec_df = pd.read_excel(data_file,sheet_name='орг-я и спец-ть СПО',dtype=str)
    org_spec_df = org_spec_df[org_spec_df['region'] == 'Республика Бурятия']
    org_spec_df['Специальность'] = org_spec_df['dr.specialty_code'] + ' ' + org_spec_df['dr.specialty_name']
    org_spec_df['cnt'] = org_spec_df['cnt'].astype(int)
    org_spec_df.drop(columns=['region'],inplace=True)
    org_spec_df['dr.short_title'] = org_spec_df['dr.short_title'].replace({'Татауровский филиал ГБПОУ "БКТИС"':'ГБПОУ "БКТИС"','Могойтинский филиал ГБПОУ "БКТИС"':'ГБПОУ "БКТИС"',
                                                                     'Усть-Баргузинский филиал ГБПОУ "БКТИС"':'ГБПОУ "БКТИС"',
                                                                     'Мухоршибирский филиал ГБПОУ "БКН"':'ГБПОУ "БКН"',
                                                                     'Кяхтинский филиал ГАПОУ "ББМК МЗ РБ"':'ГАПОУ "ББМК МЗ РБ"',
                                                                     'Хоронхойский филиал ГБПОУ "БРТСИПТ"':'ГБПОУ "БРТСИПТ"',})


    org_spec_df.rename(columns={'dr.short_title':'ПОО','dr.specialty_code':'Код','dr.specialty_name':'Наименование','cnt':'Количество поданных заявлений',
                               },inplace=True)

    # Лист дашборд СПО
    dash_spo_df = pd.read_excel(data_file, sheet_name='дашборд СПО', dtype=str)
    dash_spo_df = dash_spo_df[dash_spo_df['region_name'] == 'Республика Бурятия']

    dash_spo_df.drop(columns=['region_name','inn','kpp','okpo','entrance_test','online_application','oktmo','form_payment_code'],inplace=True)
    dash_spo_df['Специальность'] = dash_spo_df['specialty_code'] + ' ' + dash_spo_df['specialty_name']
    dash_spo_df[['Количество поданных заявлений','Вы не прошли по конкурсу (4 статус, 204)','Вы включены в приказ на зачисление (3 статус, 103)']] = dash_spo_df[['Количество поданных заявлений','Вы не прошли по конкурсу (4 статус, 204)','Вы включены в приказ на зачисление (3 статус, 103)']].astype(int)

    dash_spo_df['education_form_name'] = dash_spo_df['education_form_name'].apply(extract_educ_form)
    dash_spo_df['form_payment_title'] = dash_spo_df['form_payment_title'].apply(extract_pay_form)
    dash_spo_df['education_level_title'] = dash_spo_df['education_level_title'].apply(extract_level_educ)
    dash_spo_df['short_title'] = dash_spo_df['short_title'].replace({'Татауровский филиал ГБПОУ "БКТИС"':'ГБПОУ "БКТИС"','Могойтинский филиал ГБПОУ "БКТИС"':'ГБПОУ "БКТИС"',
                                                                     'Усть-Баргузинский филиал ГБПОУ "БКТИС"':'ГБПОУ "БКТИС"',
                                                                     'Мухоршибирский филиал ГБПОУ "БКН"':'ГБПОУ "БКН"',
                                                                     'Кяхтинский филиал ГАПОУ "ББМК МЗ РБ"':'ГАПОУ "ББМК МЗ РБ"',
                                                                     'Хоронхойский филиал ГБПОУ "БРТСИПТ"':'ГБПОУ "БРТСИПТ"',})




    dash_spo_df.rename(columns={'specialty_code':'Код','specialty_name':'Наименование','education_level_title':'Уровень образования',
                                'education_form_name':'Форма обучения','form_payment_title':'Форма оплаты','short_title':'ПОО'},inplace=True)

    # Обработка листа свод СПО
    svod_spo_df = pd.read_excel(data_file,sheet_name='свод СПО')
    svod_spo_df = svod_spo_df[svod_spo_df['Регион'] == 'Республика Бурятия']
    svod_spo_df = svod_spo_df.transpose().reset_index()
    svod_spo_df.columns = ['Показатель','Значение']
    svod_spo_df['Порядок'] = range(1,len(svod_spo_df)+1)

    # обработка листа Подано, сутки
    time_spo_df = pd.read_excel(data_file,sheet_name='Подано, сутки')

    time_spo_df = time_spo_df[time_spo_df['Регион'].isin(['day','Республика Бурятия'])]

    time_spo_df = time_spo_df.transpose().reset_index()
    time_spo_df.columns = time_spo_df.iloc[0]
    time_spo_df = time_spo_df[1:]
    time_spo_df.drop(columns=['Регион'],inplace=True)
    time_spo_df.rename(columns={'day':'Дата','Республика Бурятия':'Подано заявлений за сутки'},inplace=True)

    t = time.localtime()
    current_date = time.strftime('%d_%m_%Y', t)
    org_spec_df.to_excel(f'{end_folder}/ПОО_специальность_{current_date}.xlsx',index=False)
    dash_spo_df.to_excel(f'{end_folder}/Дашборд_СПО_{current_date}.xlsx',index=False)
    svod_spo_df.to_excel(f'{end_folder}/свод_СПО_{current_date}.xlsx',index=False)
    time_spo_df.to_excel(f'{end_folder}/подано_сутки_{current_date}.xlsx',index=False)








































if __name__ == '__main__':
    main_file = 'data/report_spo_2026.xlsx'
    main_end_folder = 'data/Результат'

    generate_data_for_dash_priem(main_file,main_end_folder)
    print('Lindy Booth')