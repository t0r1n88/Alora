"""
Скрипт для подсчета свода по районам и школам среди тех кто начал обучение и прошел хотя бы один тест
"""

import pandas as pd
import time
from tkinter import messagebox
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter


def convert_to_none(cell):
    """
    Для замены дефисов на нан
    :param cell:
    :return:
    """
    if isinstance(cell,str):
        if cell == '-':
            return None
        else:
            return cell
    else:
        return cell

def generate_svod_first_prof(list_student:str,estimation_file:str,lst_moodle:str,result_folder:str):
    """
    Функция для создания свод по прошедшим обучение.
    :param list_student: итоговый список школьников
    :param estimation_file: файл с оценками
    :param lst_moodle: файл логинами мудл
    :param result_folder: конечная папка
    :return:
    """
    try:
        # получаем время
        t = time.localtime()
        current_time = time.strftime('%H_%M на %d.%m', t)
        df = pd.read_excel(list_student,dtype=str) # файл с данными школьников
        est_df = pd.read_excel(estimation_file)
        est_df.drop(columns=['Индивидуальный номер','Учреждение (организация)','Отдел','Адрес электронной почты','Последние загруженные из этого курса'],inplace=True)

        est_df = est_df.applymap(convert_to_none)

        # Создаем колонки с ФИО
        df['ФИО'] = df['Фамилия'] + ' ' + df['Имя'] + ' '+ df['Отчество']
        est_df['ФИО'] = est_df['Фамилия'] + ' ' + est_df['Имя']

        df = pd.merge(df,est_df,how='outer',left_on='ФИО',right_on='ФИО')

        not_start_df = df[df['Тест:Тест 1.1. Речевая и логическая культура ведения делового разговора (Значение)'].isna()]
        not_start_df = not_start_df[['Школа','Класс','ФИО_представителя','Номер_телефона','Электронная_почта','ФИО','Муниципалитет']]

        # Соединяем с файлом мудла
        moodle_df = pd.read_excel(lst_moodle,dtype=str)
        moodle_df['ФИО'] = moodle_df['lastname'] + ' ' + moodle_df['firstname']

        not_start_df = pd.merge(not_start_df,moodle_df,how='inner',left_on='ФИО',right_on='ФИО')
        not_start_df.drop(columns=['firstname','lastname','email','cohort1'],inplace=True)
        not_start_df.rename(columns={'ФИО':'ФИО школьника','username':'Логин для edu-copp03.ru','password':'Пароль для edu-copp03.ru'},inplace=True)
        not_start_df['Телеграм курса'] = 'https://t.me/+Zgw_U0s4hCUzMTNi'
        not_start_df=not_start_df.reindex(columns=['ФИО_представителя','Номер_телефона','Электронная_почта','ФИО школьника','Логин для edu-copp03.ru','Пароль для edu-copp03.ru','Телеграм курса',
                                      'Муниципалитет','Школа','Класс'])

        # Сохраняем списки по муниципалитетам
        lst_value_column = not_start_df['Муниципалитет'].unique()

        for idx, value in enumerate(lst_value_column):
            wb = openpyxl.Workbook()  # создаем файл
            temp_df = not_start_df[not_start_df['Муниципалитет'] == value]  # отфильтровываем по значению
            # temp_df = temp_df[['Школа','Класс','ФИО школьника']]
            for row in dataframe_to_rows(temp_df, index=False, header=True):
                wb['Sheet'].append(row)

            # Устанавливаем автоширину для каждой колонки
            for column in wb['Sheet'].columns:
                max_length = 0
                column_name = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                wb['Sheet'].column_dimensions[column_name].width = adjusted_width

            wb.save(f'{result_folder}/{value} учеников-{len(temp_df)}.xlsx')
            wb.close()




        group_df = df.groupby(['Муниципалитет']).agg({'Тест:Тест 1.1. Речевая и логическая культура ведения делового разговора (Значение)':'count'}).fillna(0)
        group_df.rename(columns={'Тест:Тест 1.1. Речевая и логическая культура ведения делового разговора (Значение)':'Количество приступивших к обучению'},inplace=True)
        # Добавляем колонку с количеством зарегистрировавшихся
        group_df.insert(0,'Записано на курс',[6,7,12,10,7,23,32,22,20,8,30,23,34,10,10,31,8,4,38])

        group_df['% приступивших к обучению'] = round((group_df['Количество приступивших к обучению'] / group_df['Записано на курс']) * 100,1)
        group_df['Количество НЕ приступивших к обучению'] = group_df['Записано на курс'] - group_df['Количество приступивших к обучению']

        sum_row = group_df.sum(axis=0, numeric_only=True)
        sum_row = sum_row.rename('Итого').to_frame().transpose()
        group_df = pd.concat([group_df, sum_row])
        group_df.loc['Итого', '% приступивших к обучению'] = round(
            (group_df.loc['Итого', 'Количество приступивших к обучению'] / group_df.loc['Итого', 'Записано на курс']) * 100, 1)
        with pd.ExcelWriter(f'{result_folder}/Сводка Первая профессия в {current_time}.xlsx') as writer:
            group_df.to_excel(writer, sheet_name='Свод по муниципалитетам', index=True)
    except PermissionError as e:
        messagebox.showerror('Алора',
                             f'Закройте файлы созданные программой')
    except FileNotFoundError as e:
        messagebox.showerror('Алора',
                             f'Не удалось создать файл с названием {e}\n'
                             f'Уменьшите количество символов в соответствующей строке или выберите более короткий путь к итоговой папке')
    else:
        messagebox.showinfo('Алора', 'Создание документов успешно завершено !')


if __name__ == '__main__':
    main_list_student = 'data/ИТОГОВЫЙ список зарегистрировавшихся на курс.xlsx'
    main_estimation_file = 'data/Цифровой куратор Оценки.xlsx'
    main_lst_moodle = 'data/Файл для MOODLE 27_10.xlsx'
    main_result_folder = 'data/Результат'

    generate_svod_first_prof(main_list_student,main_estimation_file,main_lst_moodle,main_result_folder)

    print('Lindy Booth')