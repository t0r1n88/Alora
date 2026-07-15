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
dupl_snils = dupl_df.shape[0]

freq_stats = dupl_df['СНИЛС абитуриента'].value_counts().value_counts().sort_index()
df_freq_stats = pd.DataFrame({
    'Количество поданных заявлений': freq_stats.index,
    'Количество абитуриентов подавших указанное количество заявлений': freq_stats.values
})
df_freq_stats = df_freq_stats.sort_values(by='Количество поданных заявлений',ascending=False)

print(df_freq_stats)

# Общий свод по основным показателям
svod_df = pd.DataFrame({'Показатель':['Общее количество заявлений','Корректные СНИЛС','Некорректные СНИЛС','Уникальные СНИЛС','Повторяющиеся СНИЛС',],
                        'Значение':[all_value,correct_snils,non_correct_snils,non_dupl_snils,dupl_snils]})

# Общее количество
poo_svod_all_df = pd.pivot_table(df,index=['ПОО'],
                                 values=['СНИЛС абитуриента'],
                                 aggfunc='count',
                                 margins=True,margins_name='Итого').rename(columns={'СНИЛС абитуриента':'Количество'})

poo_svod_correct_snils_df = pd.pivot_table(snils_df,index=['ПОО'],
                                 values=['СНИЛС абитуриента'],
                                 aggfunc='count',
                                 margins=True,margins_name='Итого').rename(columns={'СНИЛС абитуриента':'Количество'})

poo_svod_non_correct_snils_df = pd.pivot_table(bad_snils_df,index=['ПОО'],
                                 values=['СНИЛС абитуриента'],
                                 aggfunc='count',
                                 margins=True,margins_name='Итого').rename(columns={'СНИЛС абитуриента':'Количество'})

poo_svod_uniq_snils_df = pd.pivot_table(snils_non_dupl_df,index=['ПОО'],
                                 values=['СНИЛС абитуриента'],
                                 aggfunc='count',
                                 margins=True,margins_name='Итого').rename(columns={'СНИЛС абитуриента':'Количество'})

poo_svod_dupl_snils_df = pd.pivot_table(dupl_df,index=['ПОО'],
                                 values=['СНИЛС абитуриента'],
                                 aggfunc='count',
                                 margins=True,margins_name='Итого').rename(columns={'СНИЛС абитуриента':'Количество'})







with pd.ExcelWriter(f'data/Свод по абитуриентам {current_time}.xlsx') as writer:
    svod_df.to_excel(writer,sheet_name='Общий свод',index=False)
    df_freq_stats.to_excel(writer,sheet_name='Свод Несколько заявлений',index=False)
    poo_svod_uniq_snils_df.to_excel(writer,sheet_name='Свод Уникальные СНИЛС',index=True)
    poo_svod_dupl_snils_df.to_excel(writer,sheet_name='Свод Повторяющиеся СНИЛС',index=True)

    poo_svod_all_df.to_excel(writer,sheet_name='Свод Все СНИЛС',index=True)
    poo_svod_correct_snils_df.to_excel(writer,sheet_name='Свод Корректные СНИЛС',index=True)
    poo_svod_non_correct_snils_df.to_excel(writer,sheet_name='Свод Некорректные СНИЛС',index=True)

    # df.to_excel(writer,sheet_name='Общий список',index=False)


print('Lindy Booth')

