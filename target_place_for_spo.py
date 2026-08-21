"""
Скрипт для подсчета количества целевиков и свода по учреждениям СПО
"""

import pandas as pd
import openpyxl
import time
import re


def check_inn(value):
    value = str(value)
    if len(value) == 9:
        return f'0{value}'
    else:
        return value



def generate_data_traget_spo(data_file:str,end_folder:str):
    """

    :param data_file: Файл из яндекс диска
    :param end_folder: Конечная папка
    """
    error_df = pd.DataFrame(columns=['Лист', 'Ошибка'])
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)

    req_wb = openpyxl.load_workbook(data_file)
    lst_sheets = req_wb.sheetnames
    lst_sheets.remove('Пример для заполнения')
    lst_cols = ['ПОО', 'Наименование работодателя', 'ИНН',
                'Количество предоставляемых целевых мест','Условия для целевиков']
    main_df = pd.DataFrame(columns=lst_cols)
    for sheet in lst_sheets:
        print(sheet)
        temp_df = pd.read_excel(data_file, sheet_name=sheet,usecols='A:D')
        temp_df.dropna(subset='ИНН',inplace=True)



        if temp_df.shape[1] == 0:
            temp_df = pd.DataFrame(columns=['Наименование работодателя', 'ИНН',
                'Количество предоставляемых целевых мест','Условия для целевиков'],
                                   data=[['Не указано','Не указано',0,'Не указано']])
        else:

            temp_df = temp_df[temp_df['Наименование работодателя'].str.strip().astype(bool)]
            temp_df = temp_df[temp_df['Наименование работодателя'].notna()]




        temp_df.insert(0,'ПОО',sheet)

        main_df = pd.concat([main_df, temp_df])


    main_df['ИНН'] = main_df['ИНН'].apply(check_inn)
    svod_df = pd.pivot_table(main_df,
                             values=['Количество предоставляемых целевых мест'],
                             index=['ПОО'],
                             aggfunc='sum',
                             fill_value=0)

    total_row = svod_df.sum(axis=0, numeric_only=True)
    total_row.name = 'Итого'  # Называем строку
    svod_df = pd.concat([svod_df, total_row.to_frame().T])
    svod_df = svod_df.reindex(
        columns=['Количество предоставляемых целевых мест'])

    main_df['ИНН'] = main_df['ИНН'].apply(lambda x:x.strip() if isinstance(x,str) else 'Отсутствует')
    not_dupl_df = main_df.drop_duplicates(subset=['ИНН'])
    dct_inn = dict(zip(not_dupl_df['ИНН'], not_dupl_df['Наименование работодателя']))

    inn_df = pd.pivot_table(main_df,
                             values=['Количество предоставляемых целевых мест'],
                             index=['ИНН'],
                             aggfunc='sum',
                             fill_value=0)

    inn_df.sort_values(by='Количество предоставляемых целевых мест',inplace=True,ascending=False)

    total_row = inn_df.sum(axis=0, numeric_only=True)
    total_row.name = 'Итого'  # Называем строку
    inn_df = pd.concat([inn_df, total_row.to_frame().T])
    inn_df = inn_df.reindex(
        columns=['Количество предоставляемых целевых мест'])

    inn_df = inn_df.reset_index()

    inn_df.columns = ['Наименование работодателя','Количество предоставляемых целевых мест']
    inn_df['Наименование работодателя'] = inn_df['Наименование работодателя'].replace(dct_inn)





    with pd.ExcelWriter(f'{end_folder}/Целевое {current_date}.xlsx') as writer:
        svod_df.to_excel(writer, sheet_name='Свод ПОО', index=True)
        inn_df.to_excel(writer, sheet_name='Свод работодатели', index=False)
        main_df.to_excel(writer, sheet_name='Общий список', index=False)








if __name__ == '__main__':
    main_data_file = 'data/Целевое обучение Работодатели.xlsx'
    main_end_result_folder = 'data/Результат'

    generate_data_traget_spo(main_data_file,main_end_result_folder)

    print('Lindy Booth')
