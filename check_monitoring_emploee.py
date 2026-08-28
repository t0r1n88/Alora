"""
Скрипт для сравнения двух папок с мониторингами
"""
import pandas as pd

import xlsxwriter
import time
import re
import os

def merge_file(folder_data:str,error_df:pd.DataFrame):
    """
    Для слияния файлов
    :param folder_data:
    :param error_df
    :return:
    """
    etalon_cols = ['1','1.1','1.2','1.3','2','3','3.1','3.2','3.3',
                   '4','4.1','4.2','4.3','5','6','7',
                   '8','9','10','11','12','13','14','15','16','17','18']
    main_cols = ['ПОО']
    main_cols.extend(etalon_cols)
    main_df = pd.DataFrame(columns=main_cols)

    for file in os.listdir(folder_data):
        print(file)
        poo = file.split('.xlsx')[0]
        temp_df = pd.read_excel(f'{folder_data}/{file}',skiprows=3)
        if len(temp_df) == 0:
            temp_error_df = pd.DataFrame(columns=['Файл', 'Ошибка'], data=[[file, 'Пустой файл']])
            error_df = pd.concat([error_df, temp_error_df])
            continue

        temp_df = temp_df.drop(columns=['Субъекты РФ',0,'0'],errors='ignore')
        temp_df.columns = etalon_cols
        temp_df.insert(0,'ПОО',poo)
        main_df = pd.concat([main_df,temp_df])

    return main_df,error_df





def check_employers(folder_data_first:str,folder_data_second:str,end_folder:str):

    error_df = error_df = pd.DataFrame(columns=['Файл', 'Ошибка'])
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)


    first_df,error_df = merge_file(folder_data_first,error_df)
    print('Обработан первый датафрейм')
    second_df, error_df = merge_file(folder_data_second,error_df)

    first_df = first_df[first_df['1'].notna()]
    first_df = first_df[~first_df['1'].str.contains('Один вариант')]
    first_df.insert(0,'Год обработки',2025)
    first_df.drop(columns=['1.1','1.2','1.3','3','3.1','3.2','3.3',
                   '4','4.1','4.2','4.3','5','6','7',
                   '8','9','10','11','12','13','14','15','16','17','18'],inplace=True)

    first_df['2'] = first_df['2'].fillna(0)
    first_df['2'] = first_df['2'].astype(int)
    first_df['ID_Объединения'] = first_df['ПОО'].astype(str) + '_' + first_df['1'].astype(str)
    first_df.columns = ['Год 2025','2025_ПОО','2025_Код','2025_Количество','ID_Объединения']


    second_df = second_df[second_df['1'].notna()]
    second_df = second_df[~second_df['1'].str.contains('Один вариант')]
    second_df.insert(0,'Год обработки',2026)
    second_df.drop(columns=['1.1','1.2','1.3','3','3.1','3.2','3.3',
                   '4','4.1','4.2','4.3','5','6','7',
                   '8','9','10','11','12','13','14','15','16','17','18'],inplace=True)

    second_df['2'] = second_df['2'].fillna(0)
    second_df['2'] = second_df['2'].astype(int)

    second_df['ID_Объединения'] = second_df['ПОО'].astype(str) + '_' + second_df['1'].astype(str)
    second_df.rename(columns={'2':'2026_Количество'},inplace=True)


    itog_df =  pd.merge(first_df, second_df, how='outer', left_on=['ID_Объединения'], right_on=['ID_Объединения'],
                           indicator=True)
    itog_df = itog_df[['Год 2025','2025_ПОО','2025_Код','2025_Количество','2026_Количество','1','ПОО','Год обработки','_merge']]
    itog_df['Проверка'] = itog_df['2025_Количество'] == itog_df['2026_Количество']

    both_error_df = itog_df[itog_df['_merge'] == 'both']

    both_error_df = both_error_df[~both_error_df['Проверка']]
    both_error_df.drop(columns=['_merge','Проверка'],inplace=True)

    left_df = itog_df[itog_df['_merge'] == 'left_only']
    left_df.drop(columns=['_merge', 'Проверка'],inplace=True)

    right_df = itog_df[itog_df['_merge'] == 'right_only']
    right_df.drop(columns=['_merge', 'Проверка'],inplace=True)

    with pd.ExcelWriter(f'{end_folder}/Сверка {current_time}.xlsx') as writer:
        both_error_df.to_excel(writer,sheet_name='Разное количество',index=False)
        left_df.to_excel(writer,sheet_name='2025',index=False)
        right_df.to_excel(writer,sheet_name='2026',index=False)










if __name__ == '__main__':
    main_first_folder = 'data/Сверка/2025'
    main_second_folder = 'data/Сверка/2026'
    main_end_folder = 'data/Сверка/Результат'

    check_employers(main_first_folder,main_second_folder,main_end_folder)

    print('Lindy Booth')

