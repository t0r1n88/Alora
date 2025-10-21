"""
Скрипт для подсчета свода по районам и школам среди тех кто начал обучение и прошел хотя бы один тест
"""

import pandas as pd
import time
from tkinter import messagebox


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

def generate_svod_first_prof(list_student:str,estimation_file:str,result_folder:str):
    """
    Функция для создания свод по прошедшим обучение.
    :param list_student: итоговый список школьников
    :param estimation_file: файл с оценками
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



        group_df = df.groupby(['Муниципалитет']).agg({'Тест:Тест 1.1. Речевая и логическая культура ведения делового разговора (Значение)':'count'}).fillna(0)
        group_df.rename(columns={'Тест:Тест 1.1. Речевая и логическая культура ведения делового разговора (Значение)':'Количество начавших обучение'},inplace=True)
        # Добавляем колонку с количеством зарегистрировавшихся
        group_df.insert(0,'Записано на курс',[6,7,13,10,7,23,32,23,21,8,30,23,34,10,10,31,8,4,39])

        group_df['% начавших обучение'] = round((group_df['Количество начавших обучение'] / group_df['Записано на курс']) * 100,1)
        group_df['Количество не начавших обучение'] = group_df['Записано на курс'] - group_df['Количество начавших обучение']

        sum_row = group_df.sum(axis=0, numeric_only=True)
        sum_row = sum_row.rename('Итого').to_frame().transpose()
        group_df = pd.concat([group_df, sum_row])
        group_df.loc['Итого', '% начавших обучение'] = round(
            (group_df.loc['Итого', 'Количество начавших обучение'] / group_df.loc['Итого', 'Записано на курс']) * 100, 1)
        with pd.ExcelWriter(f'{result_folder}/Сводка Первая профессия в {current_time}.xlsx') as writer:
            group_df.to_excel(writer, sheet_name='Свод по муниципалитетам', index=True)
    except PermissionError as e:
        messagebox.showerror('Алора',
                             f'Закройте файлы созданные программой')
    except FileNotFoundError as e:
        messagebox.showerror('Алора',
                             f'Не удалось создать файл с названием {e}\n'
                             f'Уменьшите количество символов в соответствующей строке или выберите более короткий путь к итоговой папке')
    except PackageNotFoundError as e:
        messagebox.showerror('Алора',
                             f'Не удалось создать файл с названием {e}\n'
                             f'Уменьшите количество символов в соответствующей строке файла с данными в колонке по которой создаются имена файлов или выберите более короткий путь к итоговой папке')
    else:
        messagebox.showinfo('Алора', 'Создание документов успешно завершено !')


if __name__ == '__main__':
    main_list_student = 'data/ИТОГОВЫЙ список зарегистрировавшихся на курс.xlsx'
    main_estimation_file = 'data/Цифровой куратор Оценки.xlsx'
    main_result_folder = 'data/Результат'

    generate_svod_first_prof(main_list_student,main_estimation_file,main_result_folder)

    print('Lindy Booth')