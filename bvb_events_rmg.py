"""
Скрипт для создания сводов по файлам по мероприятиям и России мои горизонты
"""

import pandas as pd
import openpyxl
import time
import re











def create_svod_bvb(bvb_data:str,rmg_data:str, end_folder:str):
    t = time.localtime()
    current_time = time.strftime('%d_%m', t)

    # Обрабатываем файл по данным Россия мои горизонты
    rmg_df = pd.read_excel(rmg_data,skiprows=2)
    rmg_df = rmg_df.iloc[1:,:]

    dct_rmg_df = dict() # словарь для хранения датафреймов по РМГ

    svod_rmg_mun_df = pd.pivot_table(rmg_df,values='Численность всех учащихся 6-11 классов (в том числе с ОВЗ и инвалидностью), зарегистрированных на платформе 2021-2025 гг.и  посетивших хотя бы одно занятие "Россия - мои горизонты"',
                                    index='Муниципалитет',
                                    aggfunc='sum',margins=True,margins_name='Итого')
    svod_rmg_mun_df.columns = ['Количество']

    dct_rmg_df['Муниципалитеты'] = svod_rmg_mun_df

    svod_rmg_school_df = pd.pivot_table(rmg_df,values='Численность всех учащихся 6-11 классов (в том числе с ОВЗ и инвалидностью), зарегистрированных на платформе 2021-2025 гг.и  посетивших хотя бы одно занятие "Россия - мои горизонты"',
                                    index=['Муниципалитет','Образовательная организация'],
                                    aggfunc='sum',margins=True,margins_name='Итого')
    svod_rmg_school_df.columns = ['Количество']

    dct_rmg_df['Школы'] = svod_rmg_school_df

    svod_rmg_class_df = pd.pivot_table(rmg_df,values='Численность всех учащихся 6-11 классов (в том числе с ОВЗ и инвалидностью), зарегистрированных на платформе 2021-2025 гг.и  посетивших хотя бы одно занятие "Россия - мои горизонты"',
                                    index=['Муниципалитет','Образовательная организация'],
                                    columns='Класс',
                                    aggfunc='sum',margins=True,margins_name='Итого')

    dct_rmg_df['Классы'] = svod_rmg_class_df

    # Создаем отдельные своды по муниципалитетам
    lst_rmg_mun = rmg_df['Муниципалитет'].unique()
    set_rmg_used_name_sheet = set() # множество для хранения названий листов

    for idx,mun in enumerate(lst_rmg_mun):
        temp_rmg_df = rmg_df[rmg_df['Муниципалитет'] == mun]
        svod_temp_rmg_df = pd.pivot_table(temp_rmg_df,values='Численность всех учащихся 6-11 классов (в том числе с ОВЗ и инвалидностью), зарегистрированных на платформе 2021-2025 гг.и  посетивших хотя бы одно занятие "Россия - мои горизонты"',
                                          index='Образовательная организация',
                                          columns='Класс',
                                          aggfunc='sum',margins=True,margins_name='Итого')
        # Укорачиваем и очищаем название
        mun = mun.replace(' муниципальный район','')

        short_value = mun[:30]  # получаем обрезанное значение
        short_value = re.sub(r'[\r\b\n\t\[\]\'+()<> :"?*|\\/]', '_', short_value)
        # проверка на случай если совпадают названия 30 символов
        if short_value not in set_rmg_used_name_sheet:
            set_rmg_used_name_sheet.add(short_value)
        else:
            short_value = f'{short_value}_{idx}'
        dct_rmg_df[short_value] = svod_temp_rmg_df











    with pd.ExcelWriter(f'{end_folder}/СВОД_РМГ_{current_time}.xlsx',engine='xlsxwriter') as writer:
        for name_sheet,out_df in dct_rmg_df.items():
            out_df.to_excel(writer,sheet_name=str(name_sheet),index=True)















if __name__ == '__main__':
    main_bvb_data = 'data/students.xlsx'
    main_rmg_data = 'data/рмг на 03.12.xlsx'
    main_end_folder = 'data/Результат'

    create_svod_bvb(main_bvb_data,main_rmg_data,main_end_folder)
    print('Lindy Booth')
