"""
Скрипт для создания сводов по файлам по мероприятиям и России мои горизонты
"""

import pandas as pd
import openpyxl
import time
import re
import os
from tkinter import messagebox






def create_events_svod(df:pd.DataFrame,name_col_events:str):
    """
    Функция для создания сводов по одному из 5 мероприятий
    """
    dct_event_df = dict()  # словарь для хранения датафреймов по мероприятию
    dct_separate = dict() # словарь для отдельных датафреймов по муниципалитетам

    # Фильтруем нулевые значения
    event_df = df[df[name_col_events] != 0]
    if len(event_df) == 0:
        return dct_event_df,dct_separate


    svod_event_mun_df = pd.pivot_table(event_df,
                                       values=name_col_events,
                                       index='Муниципалитет',
                                       aggfunc='count', margins=True, margins_name='Итого')
    svod_event_mun_df.columns = ['Количество']

    dct_event_df['Муниципалитеты'] = svod_event_mun_df

    svod_event_school_df = pd.pivot_table(event_df,
                                          values=name_col_events,
                                          index=['Муниципалитет', 'Образовательная организация'],
                                          aggfunc='count', margins=True, margins_name='Итого')
    svod_event_school_df.columns = ['Количество']

    dct_event_df['Школы'] = svod_event_school_df

    svod_event_class_df = pd.pivot_table(event_df,
                                         values=name_col_events,
                                         index=['Муниципалитет', 'Образовательная организация'],
                                         columns='Класс (без буквы)',
                                         aggfunc='count', margins=True, margins_name='Итого')

    dct_event_df['Классы'] = svod_event_class_df

    # Создаем отдельные своды по муниципалитетам
    lst_event_mun = event_df['Муниципалитет'].unique()
    set_event_used_name_sheet = set()  # множество для хранения названий листов

    for idx, mun in enumerate(lst_event_mun):
        temp_event_df = event_df[event_df['Муниципалитет'] == mun]
        svod_temp_event_df = pd.pivot_table(temp_event_df,
                                            values=name_col_events,
                                            index='Образовательная организация',
                                            columns='Класс (без буквы)',
                                            aggfunc='count', margins=True, margins_name='Итого')
        # Укорачиваем и очищаем название
        mun = mun.replace(' муниципальный район', '')

        short_value = mun[:30]  # получаем обрезанное значение
        short_value = re.sub(r'[\r\b\n\t\[\]\'+()<> :"?*|\\/]', '_', short_value)
        # проверка на случай если совпадают названия 30 символов
        if short_value not in set_event_used_name_sheet:
            set_event_used_name_sheet.add(short_value)
        else:
            short_value = f'{short_value}_{idx}'
        dct_event_df[short_value] = svod_temp_event_df
        dct_separate[short_value] = svod_temp_event_df

    return dct_event_df,dct_separate











def create_svod_bvb(bvb_data:str,rmg_data:str, end_folder:str):
    t = time.localtime()
    current_date = time.strftime('%d_%m', t)
    try:

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
        dct_separate_mun = dict() # словарь для хранения сводов по муниципалитетам для отдельных файлов

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
            dct_separate_mun[short_value] = svod_temp_rmg_df


        """
        Обработка файла мероприятиями
        """
        # Без архивных
        event_df = pd.read_excel(bvb_data,skiprows=2)
        archive_df = event_df.copy() # делаем копию вместе с архивными
        archive_df = archive_df[archive_df['Дата последнего входа на платформу'].str.contains('2025',na=False)]


        event_df = event_df[event_df['Дата архивации'] == 'Нет'] # Отбрасываем тех кто в архиве
        lst_cols_events = ['Количество пройденных диагностик (за календарный год)','Кол-во посещенных профессиональных и партнерских проб (за календарный год)',
                           'Кол-во посещенных экскурсий на предприятие (за календарный год)','Кол-во посещенных экскурсий в корпоративный музей (за календарный год)',
                           'Кол-во посещенных мастер-классов (за календарный год)']
        dct_name_file = {'Количество пройденных диагностик (за календарный год)':'Диагностики',
                         'Кол-во посещенных профессиональных и партнерских проб (за календарный год)':'Пробы',
                           'Кол-во посещенных экскурсий на предприятие (за календарный год)':'Экскурсии',
        'Кол-во посещенных экскурсий в корпоративный музей (за календарный год)':'Корпоративные музеи',
                           'Кол-во посещенных мастер-классов (за календарный год)':'Мастер-классы'}

        # Создаем папки для деления
        if not os.path.exists(f'{end_folder}/Без архивных'):
            os.makedirs(f'{end_folder}/Без архивных')

        if not os.path.exists(f'{end_folder}/С архивными'):
            os.makedirs(f'{end_folder}/С архивными')


        # без архивных
        for name_events in lst_cols_events:
            dct_name_events,dct_separate_events = create_events_svod(event_df.copy(),name_events)
            with pd.ExcelWriter(f'{end_folder}/Без архивных/БА_СВОД_{dct_name_file[name_events]}_{current_date}.xlsx', engine='xlsxwriter') as writer:
                for name_sheet, out_df in dct_name_events.items():
                    out_df.to_excel(writer, sheet_name=str(name_sheet), index=True)

            # Создаем папку для хранения отдельных по муниципалитетам
            path_events_file = f'{end_folder}/Без архивных/БА_{dct_name_file[name_events]}'  #
            if not os.path.exists(path_events_file):
                os.makedirs(path_events_file)

            for name_sheet, out_df in dct_separate_events.items():
                out_df.to_excel(f'{path_events_file}/{name_sheet}_{current_date}.xlsx', index=True)


        # с архивными

        for name_events in lst_cols_events:
            dct_name_events,dct_separate_events = create_events_svod(archive_df.copy(),name_events)
            with pd.ExcelWriter(f'{end_folder}/С архивными/А_СВОД_{dct_name_file[name_events]}_{current_date}.xlsx', engine='xlsxwriter') as writer:
                for name_sheet, out_df in dct_name_events.items():
                    out_df.to_excel(writer, sheet_name=str(name_sheet), index=True)

            # Создаем папку для хранения отдельных по муниципалитетам
            path_events_file = f'{end_folder}/С архивными/А_{dct_name_file[name_events]}'  #
            if not os.path.exists(path_events_file):
                os.makedirs(path_events_file)

            for name_sheet, out_df in dct_separate_events.items():
                out_df.to_excel(f'{path_events_file}/{name_sheet}_{current_date}.xlsx', index=True)


        """
        Сохраняем РМГ
        """
        with pd.ExcelWriter(f'{end_folder}/СВОД_РМГ_{current_date}.xlsx',engine='xlsxwriter') as writer:
            for name_sheet,out_df in dct_rmg_df.items():
                out_df.to_excel(writer,sheet_name=str(name_sheet),index=True)

        # Создаем папку для хранения отдельных по муниципалитетам
        path_rmg_file = f'{end_folder}/РМГ'  #
        if not os.path.exists(path_rmg_file):
            os.makedirs(path_rmg_file)

        for name_sheet,out_df in dct_separate_mun.items():
            out_df.to_excel(f'{path_rmg_file}/{name_sheet}_{current_date}.xlsx',index=True)


    except PermissionError as e:
        messagebox.showerror('Алора',
                             f'Закройте файлы созданные программой')
    except FileNotFoundError as e:
        messagebox.showerror('Алора',
                             f'Не удалось создать файл с названием {e}\n'
                             f'Выберите более короткий путь к конечной папке')
    else:
        messagebox.showinfo('Алора', 'Создание документов успешно завершено !')













if __name__ == '__main__':
    main_bvb_data = 'data/students.xlsx'
    main_bvb_data = 'data/Сводный У-У на 15.12.xlsx'
    main_rmg_data = 'data/рмг на 03.12.xlsx'
    main_end_folder = 'data/Результат'

    create_svod_bvb(main_bvb_data,main_rmg_data,main_end_folder)
    print('Lindy Booth')
