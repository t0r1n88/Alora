"""
Скрипт для подсчета количества уникальных абитуриентов в республике
"""

import pandas as pd

import time
import re

def clear_snils(value):
    value = str(value)

    result = re.findall(r'\d',value)
    if result:
        if len(result) == 11:
            return f'{result[0]}{result[1]}{result[2]}-{result[3]}{result[4]}{result[5]}-{result[6]}{result[7]}{result[8]} {result[9]}{result[10]}'
        else:
            return f'В СНИЛС не 11 цифр'
    else:
        return 'СНИЛС не заполнен'


def clear_fio(value):
    value = str(value)

    value_str = value.strip()
    value_str = re.sub(r'\s+',' ',value_str)
    value_str = re.sub(r'[^а-яА-ЯёЁ]','',value_str).capitalize()

    return value_str







t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
current_date = time.strftime('%d_%m_%Y', t)


df = pd.read_excel('data/Шаблон списка подавших заявления с выбором класса .xlsx',dtype=str)

df['СНИЛС абитуриента'] = df['СНИЛС абитуриента'].apply(clear_snils)

df['ФИО абитуриента'] = df['ФИО абитуриента'].apply(clear_fio)
print(df.shape)
all_value = df.shape[0] # общее количество записей

snils_df = df[~df['СНИЛС абитуриента'].isin(['В СНИЛС не 11 цифр','СНИЛС не заполнен'])]
bad_snils_df = df[df['СНИЛС абитуриента'].isin(['В СНИЛС не 11 цифр','СНИЛС не заполнен'])]
correct_snils = snils_df.shape[0] # корректные снилс
non_correct_snils = bad_snils_df.shape[0] # некорректные снилс
print(snils_df.shape)
print(bad_snils_df.shape)

snils_non_dupl_df = snils_df.drop_duplicates(subset=['СНИЛС абитуриента'])
non_dupl_snils = snils_non_dupl_df.shape[0]
print(snils_non_dupl_df.shape)
dupl_df = snils_df[snils_df['СНИЛС абитуриента'].duplicated(keep=False)]
dupl_df = dupl_df.sort_values(by='СНИЛС абитуриента')

freq_stats = dupl_df['СНИЛС абитуриента'].value_counts().value_counts().sort_index()
df_freq_stats = pd.DataFrame({
    'Количество поданных заявлений': freq_stats.index,
    'Количество абитуриентов подавших указанное количество заявлений': freq_stats.values
})
df_freq_stats = df_freq_stats.sort_values(by='Количество поданных заявлений',ascending=False)

print(df_freq_stats)
dupl_df.to_excel('data/dupl.xlsx')







df.to_excel('data/test.xlsx',index=False)


print('Lindy Booth')

