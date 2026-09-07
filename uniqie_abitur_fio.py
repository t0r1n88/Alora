"""
Скрипт для подсчета количества уникальных абитуриентов в республике по ФИО
"""

import pandas as pd
import xlsxwriter
import time
import re
import os

def clear_fio(value):
    value = str(value)

    value_str = value.strip()
    value_str = re.sub(r'\s+',' ',value_str)
    value_str = re.sub(r'[^а-яА-ЯёЁ]','',value_str).capitalize()

    return value_str


def merge_file(folder_data:str,error_df:pd.DataFrame):
    """
    Для слияния файлов
    :param folder_data:
    :param error_df
    :return:
    """
    etalon_cols = ['ФИО абитуриента','СНИЛС абитуриента','Базовое образование','Код и наименование специальности/профессии на которую подано заявление']
    main_cols = ['ПОО']
    main_cols.extend(etalon_cols)
    main_df = pd.DataFrame(columns=main_cols)

    for file in os.listdir(folder_data):
        print(file)
        poo = file.split('.xlsx')[0]
        temp_df = pd.read_excel(f'{folder_data}/{file}',dtype=str)
        if len(temp_df) == 0:
            temp_error_df = pd.DataFrame(columns=['Файл', 'Ошибка'], data=[[file, 'Пустой файл']])
            error_df = pd.concat([error_df, temp_error_df])
            continue

        diff_cols = set(etalon_cols).difference(set(temp_df.columns))
        if len(diff_cols) != 0:
            temp_error_df = pd.DataFrame(columns=['Файл', 'Ошибка'], data=[[file, f'Не хватает колонки {diff_cols}']])
            error_df = pd.concat([error_df, temp_error_df])
            continue

        temp_df = temp_df[etalon_cols]
        temp_df.insert(0,'ПОО',poo)
        main_df = pd.concat([main_df,temp_df])

    return main_df,error_df



def split_specialties_robust(text):
    if pd.isna(text):
        return []

    text = str(text).strip()

    # Ищем все коды (цифры.цифры.цифры)
    code_pattern = r'\d{2}\.\d{2}\.\d{2}'
    positions = [m.start() for m in re.finditer(code_pattern, text)]

    if not positions:
        return [text]

    result = []
    for i, pos in enumerate(positions):
        start = pos
        end = positions[i + 1] if i + 1 < len(positions) else len(text)

        part = text[start:end].strip().rstrip(',')
        if part:
            result.append(part)

    return result


def check_uniq_abitur_fio(folder_data:str,end_folder:str):

    error_df = pd.DataFrame(columns=['Файл', 'Ошибка'])
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)


    df,error_df = merge_file(folder_data,error_df)
    df = df.dropna(subset=['ФИО абитуриента'])


    # df['СНИЛС абитуриента'] = df['СНИЛС абитуриента'].apply(clear_snils)

    df['ФИО очищенное'] = df['ФИО абитуриента'].apply(clear_fio)
    # Разбиваем на списки
    # Разворачиваем списки в отдельные строки
    df['Специальности'] = df['Код и наименование специальности/профессии на которую подано заявление'].apply(split_specialties_robust)
    df = df.explode('Специальности', ignore_index=True)



    nine_df = df[df['Базовое образование'] == '9 классов'] # 9 классов
    eleven_df = df[df['Базовое образование'] == '11 классов']

    dct_df = {'Общее':df,'9 классов':nine_df,'11 классов':eleven_df}



    for name,df in dct_df.items():

        fio_non_dupl_df = df.drop_duplicates(subset=['ФИО очищенное'],keep=False)
        fio_unique_df = df.drop_duplicates(subset=['ФИО очищенное'])
        uniq_fio = fio_unique_df.shape[0]
        non_dupl_fio = fio_non_dupl_df.shape[0]
        dupl_df = df[df['ФИО очищенное'].duplicated(keep=False)]

        dupl_df = dupl_df.sort_values(by='ФИО очищенное')
        dupl_fio = dupl_df.shape[0]

        copy_dupl_df = dupl_df.copy()
        uniq_dupl_df = copy_dupl_df.drop_duplicates(subset=['ФИО очищенное'])

        dupl_uniq =  uniq_dupl_df.shape[0]
        freq_stats = dupl_df['ФИО очищенное'].value_counts().value_counts().sort_index()
        df_freq_stats = pd.DataFrame({
            'Количество поданных заявлений': freq_stats.index,
            'Количество абитуриентов подавших указанное количество заявлений': freq_stats.values
        })
        df_freq_stats = df_freq_stats.sort_values(by='Количество поданных заявлений',ascending=False)



        # Общий свод по основным показателям
        svod_df = pd.DataFrame({'Показатель':['Уникальных абитуриентов','Абитуриенты подавшие заявление на одну специальность/профессию',
                                              'Количество заявлений поданных на 2 и более специальностей/профессий','Количество абитуриентов подавших 2 и более заявлений'],
                                'Значение':[uniq_fio,non_dupl_fio,dupl_fio,dupl_uniq]})


        dct_error_snils = dict()


        lst_unique_poo = df['ПОО'].unique()

        # Подсчитываем статистику по отдельным ПОО
        main_df = pd.DataFrame(columns=['ПОО','Уникальные абитуриенты (ФИО)','Заявления на 2 и более специальностей/профессий','Количество абитуриентов подавших 2 и более заявлений'])


        for poo in lst_unique_poo:

            # Уникальные СНИЛС
            temp_fio_df = df[df['ПОО'] == poo]

            temp_uniq_fio_df =temp_fio_df.drop_duplicates(subset=['ФИО очищенное'])
            value_non_dupl = temp_uniq_fio_df.shape[0]
            # Повторяющиеся СНИЛС
            temp_dupl_df = df[df['ПОО'] == poo]
            temp_dupl_df = temp_dupl_df[temp_dupl_df['ФИО очищенное'].duplicated(keep=False)]

            value_dupl = temp_dupl_df.shape[0]


            # Подавшие заявления на одну специальность
            temp_non_dupl_snils_df = temp_fio_df.drop_duplicates(subset=['ФИО очищенное'],keep=False)
            value_non_dupl_snils = temp_non_dupl_snils_df.shape[0]

            temp_uniq_dupl_df = temp_dupl_df.drop_duplicates(subset=['ФИО очищенное'])
            value_dupl_uniq = temp_uniq_dupl_df.shape[0]

            temp_df = pd.DataFrame(columns=['ПОО','Уникальные абитуриенты (ФИО)','Абитуриенты подавшие заявление на одну специальность','Заявления на 2 и более специальностей/профессий','Количество абитуриентов подавших 2 и более заявлений'],
                                   data=[[poo,value_non_dupl,value_non_dupl_snils,value_dupl,value_dupl_uniq]])

            main_df = pd.concat([main_df,temp_df])

        main_df = main_df.sort_values(by='ПОО')
        main_df.iloc[:,1:] = main_df.iloc[:,1:].astype(int)
        # total_row = main_df.sum(axis=0)
        # total_row.name = 'Итого'  # Называем строку
        # main_df = pd.concat([main_df, total_row.to_frame().T])
        # main_df.loc['Итого','ПОО'] = 'Итого'





        with pd.ExcelWriter(f'{end_folder}/{name}_Свод по абитуриентам {current_time}.xlsx') as writer:
            svod_df.to_excel(writer,sheet_name='Общий свод',index=False)
            df_freq_stats.to_excel(writer,sheet_name='Свод Несколько заявлений',index=False)
            main_df.to_excel(writer,sheet_name='Подсчет внутри каждого ПОО',index=False)
            df.to_excel(writer,sheet_name='Общий список',index=False)
            dupl_df.to_excel(writer,sheet_name='Дубликаты',index=False)



        dct_error_snils.update({'Ошибки в структуре':error_df})
        wb = xlsxwriter.Workbook(f'{end_folder}/{name}_Ошибки {current_time}.xlsx',
                                 {'constant_memory': True, 'nan_inf_to_errors': True})
        for name_sheet, dupl_df in dct_error_snils.items():
            data_lst = dupl_df.values.tolist()  # преобразуем в список
            wb_name_sheet = wb.add_worksheet(name_sheet)  # создаем лист
            # Запись заголовков
            headers = list(dupl_df.columns)
            for col, header in enumerate(headers):
                wb_name_sheet.write(0, col, header)

            # Запись данных
            for row, data_row in enumerate(data_lst):
                for col, cell_value in enumerate(data_row):
                    wb_name_sheet.write(row + 1, col, cell_value)
        # закрываем
        wb.close()



        dupl_df.to_excel(f'{end_folder}/{name}_Дубликаты {current_time}.xlsx',index=False)
        print(error_df)

if __name__ == '__main__':
    main_data_folder = 'data/ПОО'
    main_end_folder = 'data/Результат ФИО'

    check_uniq_abitur_fio(main_data_folder,main_end_folder)
    print('Lindy Booth')