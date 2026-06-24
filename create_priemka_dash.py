"""
Скрипт для подготовки данных к загрузке в дашборд по приемной кампании
"""
import re

import pandas as pd



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

    org_spec_df.rename(columns={'dr.short_title':'ПОО','dr.specialty_code':'Код','dr.specialty_name':'Наименование','cnt':'Количество поданных заявлений',
                               },inplace=True)

    # Лист дашборд СПО
    dash_spo_df = pd.read_excel(data_file, sheet_name='дашборд СПО', dtype=str)
    # dash_spo_df = dash_spo_df[dash_spo_df['region_name'] == 'Республика Бурятия']

    dash_spo_df.drop(columns=['region_name','inn','kpp','okpo','entrance_test','online_application','oktmo'],inplace=True)
    dash_spo_df['Специальность'] = dash_spo_df['specialty_code'] + ' ' + dash_spo_df['specialty_name']
    dash_spo_df[['Количество поданных заявлений','Вы не прошли по конкурсу (4 статус, 204)','Вы включены в приказ на зачисление (3 статус, 103)']] = dash_spo_df[['Количество поданных заявлений','Вы не прошли по конкурсу (4 статус, 204)','Вы включены в приказ на зачисление (3 статус, 103)']].astype(int)

    dash_spo_df['education_form_name'] = dash_spo_df['education_form_name'].apply(extract_educ_form)




    dash_spo_df.rename(columns={'specialty_code':'Код','specialty_name':'Наименование','education_level_title':'Уровень образования',
                                'education_form_name':'Форма обучения','form_payment_title':'Оплата','short_title':'ПОО'},inplace=True)


    dash_spo_df.to_excel('data/res.xlsx')






























if __name__ == '__main__':
    main_file = 'data/report_spo_2026.xlsx'
    main_end_folder = 'data/Результат'

    generate_data_for_dash_priem(main_file,main_end_folder)
    print('Lindy Booth')