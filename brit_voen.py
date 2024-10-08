"""
Скрипт для обработки перс данных студентов БРИТ
"""
import pandas as pd
import re
import openpyxl

def extract_district(value:str):
    """
    Функция для извлечения признака района
    :param value:
    :return:
    """
    if 'Хоринск' in value:
        return 'Хоринский район'
    elif 'Улан-Удэ' in value:
        for street in dct_district['Советский']:
            if street in value:
                return 'Советский район'

        for street in dct_district['Железнодорожный']:
            if street in value:
                return 'Железнодорожный район'

        for street in dct_district['Октябрьский']:
            if street in value:
                return 'Октябрьский район'
        return 'Неопределенный район Улан-Удэ'

    elif 'район' in value:
        return 'Иной район'
    elif 'р-н' in value:
        return 'Иной район'
    else:
        return 'Неопределено'


street_df = pd.read_excel('data/Улицы.xlsx')
df = pd.read_excel('data/Студенты.xlsx')

dct_replace = {'Улица':'','улица':'','Переулок':'',}
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'улица','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'Улица','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'переулок','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'Переулок','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'Поселок','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:x.strip())

sov_df = street_df[street_df['Район'] == 'Советский']
sov_set = set(sov_df['Улица'].tolist())

okt_df = street_df[street_df['Район'] == 'Октябрьский']
okt_set = set(okt_df['Улица'].tolist())

gd_df = street_df[street_df['Район'] == 'Железнодорожный']
gd_set = set(gd_df['Улица'].tolist())

dct_district = {'Советский':sov_set,'Октябрьский':okt_set,'Железнодорожный':gd_set}

df['Адрес регистрации'] = df['Адрес регистрации'].fillna('')
df['Район'] = df['Адрес регистрации'].apply(extract_district)

df.to_excel('data/dis.xlsx',index=False)





#street_df.to_excel('data/clean_street.xlsx',index=False)