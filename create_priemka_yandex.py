"""
Скрипт для обработки яндекс таблица с результатами приемной кампании и подготовки данных для загрузки в дашборд
"""
import pandas as pd
import openpyxl









def generate_data_for_priem_yandex(data_file:str,end_folder:str):
    """

    :param data_file: Файл из яндекс диска
    :param end_folder: Конечная папка
    """
    error_df = pd.DataFrame(columns=['Лист','Ошибка'])

    req_wb = openpyxl.load_workbook(data_file)
    lst_sheets = req_wb.sheetnames
    req_wb.close()

    print(lst_sheets)
    lst_cols = ['ОУ','Код и наименование','База',
                'Финансовая основа обучения','Форма обучения','План приема',
                'КЦП','Целевые договора','План приема Профессионалитет',
                'Всего заявлений','подано через Госулуги','целевые',
                'программам Профессионалитета','участники СВО','дети участников СВО',]
    main_df = pd.DataFrame(columns=lst_cols)

    for sheet in lst_sheets:
        print(sheet)
        temp_df = pd.read_excel(data_file,sheet_name=sheet,skiprows=3,header=None)
        print(temp_df.shape)
        temp_df.dropna(how='all',inplace=True)
        # проверяем на количество
        if len(temp_df) == 0:
            temp_error_df = pd.DataFrame(columns=['Лист','Ошибка'],data=[[sheet,'На заполнен лист']])
            error_df = pd.concat([error_df,temp_error_df])
            continue

        temp_df.columns = lst_cols
        temp_df['ОУ'] = sheet
        print(temp_df.shape)




    print(error_df)










if __name__ == '__main__':
    main_data_file = 'data/ПРИЕМНАЯ КАМПАНИЯ 2026-2027.xlsx'
    main_end_result_folder = 'data/Результат'

    generate_data_for_priem_yandex(main_data_file,main_end_result_folder)

    print('Lindy Booth')

