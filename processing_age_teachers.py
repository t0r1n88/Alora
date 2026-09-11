"""
Скрипт для обработки данных по распределению по возрастам преподавателей и мастеров производственного обучения
"""

import pandas as pd
import openpyxl
import time
import re
from tkinter import messagebox


class WrongOrderRow(Exception):
    """
    Для отслеживания неправильного порядка строк с категориями
    """
    pass



def clean_number(value):
    if pd.isna(value):
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    try:
        return int(value)
    except:
        return 0



def processing_age_teachers(data_file:str,end_folder:str):
    try:
        error_df = pd.DataFrame(columns=['Лист', 'Ошибка'])
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        current_date = time.strftime('%d_%m_%Y', t)

        req_wb = openpyxl.load_workbook(data_file)
        lst_sheets = req_wb.sheetnames
        lst_cols = ['Категория','моложе 25 лет','25-29 лет','30-34 лет',
                    '35-39 лет','40-44 лет','45-49 лет',
                    '50-54 лет','55-59 лет','60-64 лет','65 лет и более'
                    ]
        # без колонки категории чтобы привести к инту
        lst_not_cat = ['моложе 25 лет','25-29 лет','30-34 лет',
                    '35-39 лет','40-44 лет','45-49 лет',
                    '50-54 лет','55-59 лет','60-64 лет','65 лет и более'
                    ]
        lst_cat = ['преподаватели спец дисциплин','мастера производственного обучения','преподаватели общеобразовательных предметов всего:',
                   'русского языка и литературы','языка народов России и литературы','истории, экономики, права, обществознания ',
                   'информатики','физики','математики',
                   'химии','географии','биологии',
                   'иностранных языков всего:','английского языка','немецкого языка',
                   'французского языка','китайского языка','физической культуры',
                   'труда (технологии)','музыки и пения','изобразительного искусства, черчения',
                   'основ безопасности и защиты Родины','основ духовно-нравственной культуры народов России','прочих предметов'
                   ]

        main_df = pd.DataFrame(columns=['ПОО'].extend(lst_cols))

        lst_df = [] # для хранения

        for sheet in lst_sheets:
            print(sheet)
            temp_df = pd.read_excel(data_file, sheet_name=sheet)
            temp_df = temp_df.dropna(axis=1, how='all')

            try:
                temp_df = temp_df[lst_cols]
                # проверяем правильность строк
                temp_cat = temp_df['Категория'].tolist()
                if lst_cat != temp_cat:
                    diff_cat =list(set(temp_cat).difference(set(lst_cat)))
                    if len(diff_cat) == 0:
                        for idx,cat in enumerate(lst_cat):
                            if cat != temp_cat[idx]:
                                diff_cat.append(f'{cat} не равно {temp_cat[idx]}')

                    raise WrongOrderRow

                sheet_df = temp_df.copy()
                sheet_df[lst_not_cat] = sheet_df[lst_not_cat].applymap(clean_number)
                # собираем
                lst_df.append(sheet_df)

                temp_df = sheet_df.copy()
                temp_df.insert(0,'ПОО',sheet)
                main_df = pd.concat([main_df,temp_df])

            except KeyError as e:
                temp_error_df = pd.DataFrame(columns=['Лист', 'Ошибка'], data=[[sheet, f' не найдена колонка{e.args}']])
                error_df = pd.concat([error_df, temp_error_df])
                continue
            except WrongOrderRow:
                temp_error_df = pd.DataFrame(columns=['Лист', 'Ошибка'], data=[[sheet, f'отличаются строки или порядок строк {diff_cat}']])
                error_df = pd.concat([error_df, temp_error_df])
                continue

        # Объединяем и суммируем
        combined = pd.concat(lst_df, ignore_index=True)
        result = combined.groupby('Категория', as_index=False).sum()

        # Восстанавливаем порядок категорий
        result['Категория'] = pd.Categorical(result['Категория'],
                                             categories=lst_cat,
                                             ordered=True)
        result = result.sort_values('Категория').reset_index(drop=True)

        # Добавляем итоговую строку
        total = result.sum(numeric_only=True)
        total['Категория'] = 'ИТОГО'
        result = pd.concat([result, pd.DataFrame([total])], ignore_index=True)

        result.insert(1,'Всего',result[lst_not_cat].sum(axis=1))

        result.insert(5,'Всего моложе 25-34 лет',result[['моложе 25 лет','25-29 лет','30-34 лет']].sum(axis=1))
        result.insert(9,'Всего 35-49 лет',result[['35-39 лет','40-44 лет','45-49 лет']].sum(axis=1))
        result.insert(14,'Всего 50-65 лет и более',result[['50-54 лет','55-59 лет','60-64 лет','65 лет и более']].sum(axis=1))

        dct_row = {'преподаватели спец дисциплин':'Спецдисциплины','мастера производственного обучения':'Мастера','преподаватели общеобразовательных предметов всего:':'Общие дисциплины',
                   'русского языка и литературы':'Русский,литература','языка народов России и литературы':'язык народов России',
                   'истории, экономики, права, обществознания ':'История и т.п.','информатики':'Информатика',
                   'физики':'Физика','математики':'Математика',
                   'химии':'Химия','географии':'География',
                   'биологии':'Биология','иностранных языков всего:':'Иняз Всего',
                   'английского языка':'Английский','немецкого языка':'Немецкий',
                   'французского языка':'Французкий','китайского языка':'Китайский',
                   'физической культуры':'Физ-ра','труда (технологии)':'Труд',
                   'музыки и пения':'Музыка,пение','изобразительного искусства, черчения':'ИЗО',
                   'основ безопасности и защиты Родины':'ОБЖ','прочих предметов':'прочее',
                   'основ духовно-нравственной культуры народов России':'ДНК'
                   }

        cat_dct = dict()

        for row, name_sheet in dct_row.items():
            temp_df = main_df[main_df['Категория'] == row]
            temp_df = temp_df.sort_values(by='ПОО')
            temp_df.insert(2, 'Всего', temp_df[lst_not_cat].sum(axis=1))
            temp_df.insert(6, 'Всего моложе 25-34 лет', temp_df[['моложе 25 лет', '25-29 лет', '30-34 лет']].sum(axis=1))
            temp_df.insert(10, 'Всего 35-49 лет', temp_df[['35-39 лет', '40-44 лет', '45-49 лет']].sum(axis=1))
            temp_df.insert(15, 'Всего 50-65 лет и более',
                          temp_df[['50-54 лет', '55-59 лет', '60-64 лет', '65 лет и более']].sum(axis=1))
            cat_dct[name_sheet] = temp_df


        out_dct = {'Свод':result,'Список':main_df}
        out_dct.update(cat_dct)

        with pd.ExcelWriter(f'{end_folder}/Результат {current_time}.xlsx') as writer:
            for name, df in out_dct.items():
                df.to_excel(writer,sheet_name=name,index=False)


        print(error_df)

        if len(error_df) != 0:
            error_df.to_excel(f'{end_folder}/Ошибки {current_time}.xlsx',index=False)


    except PermissionError as e:
        messagebox.showerror('Пенни',
                             f'Закройте файлы созданные программой')
    except FileNotFoundError as e:
        messagebox.showerror('Пенни',
                             f'Не удалось создать файл с названием {e}\n'
                             f'Выберите более короткий путь к конечной папке')
    else:
        messagebox.showinfo('Пенни', 'Создание документов успешно завершено !')







if __name__ == '__main__':
    main_data_file = 'data/Распределение преподавателей и мастеров производственного обучения по возрасту.xlsx'
    main_end_folder = 'data/Результат Возраст'
    processing_age_teachers(main_data_file,main_end_folder)

    print('Lindy Booth')

