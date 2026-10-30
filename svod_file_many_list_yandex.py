"""
Скрипт для сбора в одну таблицу данных из яндекс таблицы с многими листами
"""

import pandas as pd
import openpyxl
import time
import re


def generate_data_yandex_file(data_file:str,end_folder:str):
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

    lst_cols = ['Муниципалитет','Количество первокурсников']
    lst_main_cols = ['ПОО','Муниципалитет','Количество первокурсников']

    main_df = pd.DataFrame(columns=lst_cols)
    for sheet in lst_sheets:
        print(sheet)
        temp_df = pd.read_excel(data_file, sheet_name=sheet)
        # проверяем наличие нужных колонок
        diff_cols = set(lst_cols).difference(set(temp_df.columns))

        if len(diff_cols) !=0 :
            temp_error_df = pd.DataFrame(columns=['Лист', 'Ошибка'], data=[[sheet, f'{diff_cols}']])
            error_df = pd.concat([error_df, temp_error_df])
            continue

        temp_df = temp_df[lst_cols]
        temp_df = temp_df.dropna(subset=['Муниципалитет'])
        temp_df = temp_df[~temp_df['Муниципалитет'].str.contains('Итого|Всего', case=False, na=False)]
        temp_df['Муниципалитет'] = temp_df['Муниципалитет'].apply(lambda x:x.strip() if isinstance(x,str) else x)
        temp_df.insert(0,'ПОО',sheet)

        main_df = pd.concat([main_df, temp_df])



    print(error_df)
    # Проверочный свод
    poo_svod_df = pd.pivot_table(main_df,
                             values=['Количество первокурсников'],
                             index=['ПОО'],
                             aggfunc='sum',
                             fill_value=0,)

    total_row = poo_svod_df.sum(axis=0, numeric_only=True)
    total_row.name = 'Итого'  # Называем строку
    poo_svod_df = pd.concat([poo_svod_df, total_row.to_frame().T])

    mun_svod_df = pd.pivot_table(main_df,
                             values=['Количество первокурсников'],
                             index=['Муниципалитет'],
                             aggfunc='sum',
                             fill_value=0,)
    total_sum = mun_svod_df['Количество первокурсников'].sum()

    mun_svod_df['% от общего'] =(mun_svod_df['Количество первокурсников'] / total_sum) * 100
    mun_svod_df['% от общего'] = (mun_svod_df['% от общего']).map('{:.2f}%'.format)


    total_row = mun_svod_df.sum(axis=0, numeric_only=True)
    total_row.name = 'Итого'  # Называем строку
    mun_svod_df = pd.concat([mun_svod_df, total_row.to_frame().T])

    mun_poo_svod_df = pd.pivot_table(main_df,
                             values=['Количество первокурсников'],
                             index=['Муниципалитет'],
                            columns=['ПОО'],
                             aggfunc='sum',
                             fill_value=0,)

    total_row = mun_poo_svod_df.sum(axis=0, numeric_only=True)
    total_row.name = 'Итого'  # Называем строку
    mun_poo_svod_df = pd.concat([mun_poo_svod_df, total_row.to_frame().T])





    with pd.ExcelWriter(f'{end_folder}/Итоговый файл {current_time}.xlsx') as writer:
        mun_svod_df.to_excel(writer, sheet_name='Свод по муниципалитетам', index=True)
        mun_poo_svod_df.to_excel(writer, sheet_name='Свод по муниципалитетам и ПОО', index=True)
        poo_svod_df.to_excel(writer, sheet_name='ПОО заполнение', index=True)
        main_df.to_excel(writer, sheet_name='Общий список', index=False)









if __name__ == '__main__':
    main_data_file = 'data/Откуда прибыли первокурсники.xlsx'
    main_end_folder = 'data/Результат'

    generate_data_yandex_file(main_data_file,main_end_folder)

    print('Lindy Booth')






