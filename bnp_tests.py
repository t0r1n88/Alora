"""
Скрипт для обработки результатов тестирования конкурса Лучший по профессии Бурятнефтепродукт
"""

import pandas as pd
import warnings
warnings.filterwarnings('ignore', category=pd.errors.SettingWithCopyWarning)
from tkinter import messagebox









def generate_result_bnp(file_data:str,end_folder:str):
    try:
        df = pd.read_excel(file_data)

        lst_unique = df['Выберите свою номинацию'].unique()

        dct_value = {'Оператор заправщик':2,
                     'Водитель Бензовоза':122,
                     'Работник кафе':242,
                     'Лаборант химического анализа':362,
                     'Оператор товарный':482,
                     'Оператор кассир':602,
                     'Старший смены':722}


        for cat,value in dct_value.items():
            print(cat)
            temp_df = df[df['Выберите свою номинацию'] == cat]
            person_df = temp_df.iloc[:,0:2]
            answers_df = temp_df.iloc[:,value:value + 120]
            answers_df = pd.concat([person_df,answers_df],axis=1)

            # Список колонок с баллами
            lst_cols_ball = [column for column in answers_df.columns if '/ Баллы' in column]
            answers_df[lst_cols_ball] = answers_df[lst_cols_ball].astype(int,errors='ignore') # делаем интовыми
            answers_df[lst_cols_ball] = answers_df[lst_cols_ball].replace(1, 0.5) # заменяем чтобы не было путаницы





            # считаем общую сумму
            answers_df['Общая сумма баллов'] = answers_df[lst_cols_ball].sum(axis=1)
            person_df['Общая сумма баллов'] = answers_df[lst_cols_ball].sum(axis=1)


            # считаем сумму первых 45
            answers_df['Баллы по профессии'] = answers_df[lst_cols_ball[:45]].sum(axis=1)
            person_df['Баллы по профессии'] = answers_df[lst_cols_ball[:45]].sum(axis=1)
            # считаем баллы по
            answers_df['Баллы по ПБОТС'] = answers_df[lst_cols_ball[45:]].sum(axis=1)

            person_df['Баллы по ПБОТС'] = answers_df[lst_cols_ball[45:]].sum(axis=1)

            # Сортируем
            answers_df.sort_values(by='Общая сумма баллов',ascending=False,inplace=True)
            person_df.sort_values(by='Общая сумма баллов',ascending=False,inplace=True)

            with pd.ExcelWriter(f'{end_folder}/{cat}.xlsx') as writer:
                person_df.to_excel(writer,sheet_name='Краткие результаты',index=False)
                answers_df.to_excel(writer,sheet_name='Полные результаты',index=False)
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

    main_data = 'data/data.xlsx'
    main_end_folder = 'data/Результат'

    generate_result_bnp(main_data,main_end_folder)

    print('Lindy Booth')




