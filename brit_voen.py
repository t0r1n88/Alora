"""
Скрипт для обработки перс данных студентов БРИТ
"""
import pandas as pd
import re
import openpyxl

street_df = pd.read_excel('data/Улицы.xlsx')
dct_replace = {'Улица':'','улица':'','Переулок':'',}
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'улица','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'Улица','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'переулок','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'Переулок','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:re.sub(r'Поселок','',x))
street_df['Улица'] = street_df['Улица'].apply(lambda x:x.strip())

print(street_df)
street_df.to_excel('data/clean_street.xlsx',index=False)