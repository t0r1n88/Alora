"""
Скрипт для подсчета количества уникальных абитуриентов в республике
"""

import pandas as pd

import time
import re
import os

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



def merge_file(folder_data:str,error_df:pd.DataFrame):
    """
    Для слияния файлов
    :param folder_data:
    :param error_df
    :return:
    """
    etalon_cols = ['ФИО абитуриента','СНИЛС абитуриента','Базовое образование','Код и наименование специальности/профессии на которую подано заявление']
    main_cols = ['ПОО']
    main_cols.extend(etalon_cols)
    main_df = pd.DataFrame(columns=main_cols)

    for file in os.listdir(folder_data):
        print(file)
        poo = file.split('.xlsx')[0]
        temp_df = pd.read_excel(f'{folder_data}/{file}',dtype=str)
        if len(temp_df) == 0:
            temp_error_df = pd.DataFrame(columns=['Файл', 'Ошибка'], data=[[file, 'Пустой файл']])
            error_df = pd.concat([error_df, temp_error_df])
            continue

        diff_cols = set(etalon_cols).difference(set(temp_df.columns))
        if len(diff_cols) != 0:
            temp_error_df = pd.DataFrame(columns=['Файл', 'Ошибка'], data=[[file, f'Не хватает колонки {diff_cols}']])
            error_df = pd.concat([error_df, temp_error_df])
            continue

        temp_df = temp_df[etalon_cols]
        temp_df.insert(0,'ПОО',poo)
        main_df = pd.concat([main_df,temp_df])

    return main_df,error_df


def split_specialties_robust(text):
    if pd.isna(text):
        return []

    text = str(text).strip()

    # Ищем все коды (цифры.цифры.цифры)
    code_pattern = r'\d{2}\.\d{2}\.\d{2}'
    positions = [m.start() for m in re.finditer(code_pattern, text)]

    if not positions:
        return [text]

    result = []
    for i, pos in enumerate(positions):
        start = pos
        end = positions[i + 1] if i + 1 < len(positions) else len(text)

        part = text[start:end].strip().rstrip(',')
        if part:
            result.append(part)

    return result





def check_uniq_abitur(folder_data:str,end_folder:str):

    error_df = error_df = pd.DataFrame(columns=['Файл', 'Ошибка'])
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)


    df,error_df = merge_file(folder_data,error_df)
    df = df.dropna(subset=['СНИЛС абитуриента'])


    df['СНИЛС абитуриента'] = df['СНИЛС абитуриента'].apply(clear_snils)

    # df['ФИО абитуриента'] = df['ФИО абитуриента'].apply(clear_fio)
    # Разбиваем на списки
    # Разворачиваем списки в отдельные строки
    df['Специальности'] = df['Код и наименование специальности/профессии на которую подано заявление'].apply(split_specialties_robust)
    df = df.explode('Специальности', ignore_index=True)


    all_value = df.shape[0] # общее количество записей

    snils_df = df[~df['СНИЛС абитуриента'].isin(['В СНИЛС не 11 цифр','СНИЛС не заполнен'])]
    bad_snils_df = df[df['СНИЛС абитуриента'].isin(['В СНИЛС не 11 цифр','СНИЛС не заполнен'])]
    correct_snils = snils_df.shape[0] # корректные снилс
    non_correct_snils = bad_snils_df.shape[0] # некорректные снилс

    snils_non_dupl_df = snils_df.drop_duplicates(subset=['СНИЛС абитуриента'],keep=False)
    snils_unique_df = snils_df.drop_duplicates(subset=['СНИЛС абитуриента'])
    uniq_snils = snils_unique_df.shape[0]
    non_dupl_snils = snils_non_dupl_df.shape[0]
    dupl_df = snils_df[snils_df['СНИЛС абитуриента'].duplicated(keep=False)]
    dupl_df = dupl_df.sort_values(by='СНИЛС абитуриента')
    dupl_snils = dupl_df.shape[0]

    uniq_dupl_df = dupl_df.drop_duplicates(subset=['СНИЛС абитуриента'])
    dupl_uniq =  uniq_dupl_df.shape[0]

    freq_stats = dupl_df['СНИЛС абитуриента'].value_counts().value_counts().sort_index()
    df_freq_stats = pd.DataFrame({
        'Количество поданных заявлений': freq_stats.index,
        'Количество абитуриентов подавших указанное количество заявлений': freq_stats.values
    })
    df_freq_stats = df_freq_stats.sort_values(by='Количество поданных заявлений',ascending=False)


    # Общий свод по основным показателям
    svod_df = pd.DataFrame({'Показатель':['Общее количество заявлений','Корректные СНИЛС','Некорректные СНИЛС','Уникальных абитуриентов','Абитуриенты подавшие заявление на одну специальность/профессию',
                                          'Количество заявлений поданных на 2 и более специальностей/профессий','Количество абитуриентов подавших 2 и более заявлений'],
                            'Значение':[all_value,correct_snils,non_correct_snils,uniq_snils,non_dupl_snils,dupl_snils,dupl_uniq]})

    # Общее количество
    poo_svod_all_df = pd.pivot_table(df,index=['ПОО'],
                                     values=['СНИЛС абитуриента'],
                                     aggfunc='count',
                                     margins=True,margins_name='Итого').rename(columns={'СНИЛС абитуриента':'Количество'})


    lst_unique_poo = df['ПОО'].unique()

    # Подсчитываем статистику по отдельным ПОО
    main_df = pd.DataFrame(columns=['ПОО','Корректные СНИЛС','Некорректные СНИЛС','Уникальные абитуриенты (СНИЛС)','Заявления на 2 и более специальностей/профессий'])


    for poo in lst_unique_poo:
        # Корректные СНИЛС
        temp_snils_df = snils_df[snils_df['ПОО'] == poo]
        value_correct_snils = temp_snils_df.shape[0]
        # Некорректные снилс
        temp_bad_snils_df = bad_snils_df[bad_snils_df['ПОО'] == poo]
        value_bad_snils = temp_bad_snils_df.shape[0]

        # Уникальные СНИЛС
        temp_non_dupl_df = snils_unique_df[snils_unique_df['ПОО'] == poo]
        value_non_dupl = temp_non_dupl_df.shape[0]
        # Повторяющиеся СНИЛС
        temp_dupl_df = dupl_df[dupl_df['ПОО'] == poo]
        value_dupl = temp_dupl_df.shape[0]

        temp_df = pd.DataFrame(columns=['ПОО','Корректные СНИЛС','Некорректные СНИЛС','Уникальные абитуриенты (СНИЛС)','Заявления на 2 и более специальностей/профессий'],
                               data=[[poo,value_correct_snils,value_bad_snils,value_non_dupl,value_dupl]])

        main_df = pd.concat([main_df,temp_df])

    main_df = main_df.sort_values(by='ПОО')
    main_df.iloc[:,1:] = main_df.iloc[:,1:].astype(int)
    total_row = main_df.sum(axis=0)
    total_row.name = 'Итого'  # Называем строку
    main_df = pd.concat([main_df, total_row.to_frame().T])
    main_df.loc['Итого','ПОО'] = 'Итого'





    with pd.ExcelWriter(f'{end_folder}/Свод по абитуриентам {current_time}.xlsx') as writer:
        svod_df.to_excel(writer,sheet_name='Общий свод',index=False)
        df_freq_stats.to_excel(writer,sheet_name='Свод Несколько заявлений',index=False)
        main_df.to_excel(writer,sheet_name='Свод по ПОО',index=False)


        # df.to_excel(writer,sheet_name='Общий список',index=False)
    error_df.to_excel(f'{end_folder}/Ошибки {current_time}.xlsx',index=False)
    dupl_df.to_excel(f'{end_folder}/Дубликаты {current_time}.xlsx',index=False)
    print(error_df)

if __name__ == '__main__':
    main_data_folder = 'data/ПОО'
    main_end_folder = 'data'

    check_uniq_abitur(main_data_folder,main_end_folder)

    print('Lindy Booth')

