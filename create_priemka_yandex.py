"""
Скрипт для обработки яндекс таблица с результатами приемной кампании и подготовки данных для загрузки в дашборд
"""
import pandas as pd
import openpyxl
import time
import re
from tkinter import messagebox




def extract_educ_form(value):
    if pd.isna(value):
        return 'Не заполнена форма обучения'

    form_str = str(value).strip()

    if re.search(r'\bочная\b',form_str,re.IGNORECASE):
        return 'очная'
    elif re.search(r'\bОчно\b',form_str):
        return 'очная'
    elif re.search(r'\очно[- ]заочн|очн[оы]е?[- ]заочн|оч[-.]заоч',form_str,re.IGNORECASE):
        return 'очно-заочная'
    elif re.search(r'\bзаочная\b',form_str,re.IGNORECASE):
        return 'заочная'

    else:
        return f'{value} неизвестная форма обучения'


def extract_pay_form(value):
    if pd.isna(value):
        return 'Не заполнена форма оплаты'

    form_str = str(value).strip()

    if re.search(r'внебюджет',form_str,re.IGNORECASE):
        return 'коммерческая'
    elif re.search(r'бюджет',form_str,re.IGNORECASE):
        return 'бюджет'
    elif re.search(r'коммер|платн|оплат|договор|ком',form_str,re.IGNORECASE):
        return 'коммерческая'

    else:
        return f'{value} неизвестная форма оплаты'


def extract_level_educ(value):
    if pd.isna(value):
        return 'ОВЗ или Не заполнена База'

    form_str = str(value).strip()

    if re.search(r'9',form_str,re.IGNORECASE):
        return '9 класс'
    elif re.search(r'11',form_str,re.IGNORECASE):
        return '11 класс'
    elif re.search(r'4',form_str,re.IGNORECASE):
        return '4 класс'
    elif re.search(r'основное',form_str,re.IGNORECASE):
        return '9 класс'
    elif re.search(r'среднее',form_str,re.IGNORECASE):
        return '11 класс'
    elif re.search(r'ООО',form_str,re.IGNORECASE):
        return '9 класс'
    elif re.search(r'СОО',form_str,re.IGNORECASE):
        return '11 класс'

    else:
        return f'{value} неизвестный уровень образования'







def generate_data_for_priem_yandex(data_file:str,end_folder:str):
    """

    :param data_file: Файл из яндекс диска
    :param end_folder: Конечная папка
    """
    try:
        error_df = pd.DataFrame(columns=['Лист','Ошибка'])
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        current_date = time.strftime('%d_%m_%Y', t)

        req_wb = openpyxl.load_workbook(data_file)
        lst_sheets = req_wb.sheetnames
        req_wb.close()
        lst_cols = ['ОУ','Код и наименование','База',
                    'Финансовая основа обучения','Форма обучения','План приема',
                    'КЦП','Целевые договора','План приема Профессионалитет',
                    'Всего заявлений','подано Госулуги','подано целевые',
                    'подано Профессионалитет','участники СВО','дети участников СВО',]
        main_df = pd.DataFrame(columns=lst_cols)

        for sheet in lst_sheets:
            print(sheet)
            temp_df = pd.read_excel(data_file,sheet_name=sheet,skiprows=3,header=None)
            # проверяем на количество колонок
            if len(temp_df.columns) % 3 != 0:
                temp_error_df = pd.DataFrame(columns=['Лист','Ошибка'],data=[[sheet,'Количество колонок не кратно 3']])
                error_df = pd.concat([error_df,temp_error_df])
            cols_to_keep = (len(temp_df.columns) // 3) * 3  # Целочисленное деление на 3
            # Оставляем только первые cols_to_keep колонок
            temp_df = temp_df.iloc[:, :cols_to_keep]

            temp_df = pd.concat([
                temp_df.iloc[:, :9],  # первые 9 колонок
                temp_df.iloc[:, -6:]  # последние 6 колонок
            ], axis=1)
            temp_df.dropna(inplace=True,thresh=5)
            # проверяем на количество
            if len(temp_df) == 0:
                temp_error_df = pd.DataFrame(columns=['Лист','Ошибка'],data=[[sheet,'На заполнен лист']])
                error_df = pd.concat([error_df,temp_error_df])
                continue

            temp_df.columns = lst_cols
            temp_df['ОУ'] = sheet
            temp_df = temp_df[temp_df['Код и наименование'].notna()] # отбрасываем строки у которых не записан Код специальности
            if len(temp_df) == 0:
                temp_error_df = pd.DataFrame(columns=['Лист','Ошибка'],data=[[sheet,'На заполнен лист']])
                error_df = pd.concat([error_df,temp_error_df])
                continue
            pattern = '|'.join(['итог','всего'])
            temp_df = temp_df[~temp_df['Код и наименование'].str.contains(pattern,case=False, na=False)]

            temp_df = temp_df.applymap(
                lambda x: re.sub(r'\s+', ' ', x) if isinstance(x, str) else x)  # очищаем от лишних пробелов
            temp_df = temp_df.applymap(
                lambda x: x.strip() if isinstance(x, str) else x)  # очищаем от пробелов в начале и конце

            main_df = pd.concat([main_df,temp_df])

        main_df['Форма обучения'] = main_df['Форма обучения'].apply(extract_educ_form)
        main_df['Финансовая основа обучения'] = main_df['Финансовая основа обучения'].apply(extract_pay_form)
        main_df['База'] = main_df['База'].apply(extract_level_educ)
        main_df.sort_values(by='ОУ',inplace=True)
        main_df.fillna(0,inplace=True)
        main_df[main_df.columns[-10:]] = main_df[main_df.columns[-10:]].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)


        svod_df = pd.pivot_table(main_df,
                                 values=['Всего заявлений','подано Госулуги','подано целевые','подано Профессионалитет','участники СВО','дети участников СВО'],
                                 index=['ОУ'],
                                 aggfunc='sum',
                                 fill_value=0,)

        total_row = svod_df.sum(axis=0, numeric_only=True)
        total_row.name = 'Итого'  # Называем строку
        svod_df = pd.concat([svod_df, total_row.to_frame().T])
        svod_df = svod_df.reindex(columns=['Всего заявлений','подано Госулуги','подано Профессионалитет','подано целевые','участники СВО','дети участников СВО'])






        with pd.ExcelWriter(f'{end_folder}/Итоговый файл {current_date}.xlsx') as writer:
            svod_df.to_excel(writer,sheet_name='Свод',index=True)
            main_df.to_excel(writer,sheet_name='Общий список',index=False)

        print(error_df)
        error_df.to_excel(f'{end_folder}/Ошибки_{current_time}.xlsx',index=False)

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
    main_data_file = 'data/ПРИЕМНАЯ КАМПАНИЯ 2026-2027.xlsx'
    main_end_result_folder = 'data/Результат'

    generate_data_for_priem_yandex(main_data_file,main_end_result_folder)

    print('Lindy Booth')

