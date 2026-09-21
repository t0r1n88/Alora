"""
Скрипт для подсчета данных по трудоустройству по УГС
"""

import pandas as pd
import openpyxl
import time
import re


def find_match(value, search_list):
    """Ищет первое значение из search_list, входящее в value"""
    if pd.isna(value):
        return None
    # Нормализация: убираем всё лишнее, приводим к нижнему регистру
    value_clean = re.sub(r'\s+', ' ', str(value).replace('\xa0', ' ')).strip().lower()

    if not value_clean:
        return None

    for item in search_list:
        if pd.isna(item):
            continue
        item_clean = re.sub(r'\s+', ' ', str(item).replace('\xa0', ' ')).strip().lower()

        # Ищем в обе стороны
        if value_clean in item_clean:
            return item

    return None

def extract_ugs(value):
    if pd.isna(value):
        return 'Не заполнен Код и наименование'

    value = str(value).strip()

    if value.startswith('08.'):
        return '08.00.00 Техника и технологии строительства'
    elif value.startswith('09.'):
        return
    elif value.startswith('10.'):
        return '10.00.00 Информационная безопасность'
    elif value.startswith('11.'):
        return '11.00.00 Электроника, радиотехника и системы связи'
    elif value.startswith('12.'):
        return '12.00.00 Фотоника, приборостроение, оптические и биотехнические системы и технологии'
    elif value.startswith('13.'):
        return '13.00.00 Электро- и теплоэнергетика'
    elif value.startswith('15.'):
        return '15.00.00 Машиностроение'
    elif value.startswith('18.'):
        return '18.00.00 Химические технологии'
    elif value.startswith('19.'):
        return '19.00.00 Промышленная экология и биотехнологии'
    elif value.startswith('20.'):
        return '20.00.00 Техносферная безопасность и природообустройство'
    elif value.startswith('21.'):
        return '21.00.00 Прикладная геология, горное дело, нефтегазовое дело и геодезия'
    elif value.startswith('23.'):
        return '23.00.00 Техника и технологии наземного транспорта'
    elif value.startswith('24.'):
        return '24.00.00 Авиационная и ракетно-космическая техника'
    elif value.startswith('25.'):
        return '25.00.00 Аэронавигация и эксплуатация авиационной и ракетно-космической техники'
    elif value.startswith('27.'):
        return '27.00.00 Управление в технических системах'
    elif value.startswith('29.'):
        return '29.00.00 Технологии легкой промышленности'
    elif value.startswith('31.'):
        return '31.00.00 Клиническая медицина'
    elif value.startswith('33.'):
        return '33.00.00 Фармация'
    elif value.startswith('34.'):
        return '34.00.00 Сестринское дело'
    elif value.startswith('35.'):
        return '35.00.00 Сельское, лесное и рыбное хозяйство'
    elif value.startswith('36.'):
        return '36.00.00 Ветеринария и зоотехния'
    elif value.startswith('38.'):
        return '38.00.00 Экономика и управление'
    elif value.startswith('39.'):
        return '39.00.00 Социология и социальная работа'
    elif value.startswith('40.'):
        return '40.00.00 Юриспруденция'
    elif value.startswith('43.'):
        return '43.00.00 Сервис и туризм'
    elif value.startswith('44.'):
        return '44.00.00 Образование и педагогические науки'
    elif value.startswith('46.'):
        return '46.00.00 Гуманитарные науки'
    elif value.startswith('49.'):
        return '49.00.00 Физическая культура и спорт'
    elif value.startswith('51.'):
        return '51.00.00 Культуроведение и социокультурные проекты'
    elif value.startswith('52.'):
        return '52.00.00 Сценические искусства и литературное творчество'
    elif value.startswith('53.'):
        return '53.00.00 Музыкальное искусство'
    elif value.startswith('54.'):
        return '54.00.00 Изобразительное и прикладные виды искусств'
    elif value.startswith('55.'):
        return '55.00.00 Экранные искусства'
    elif re.search(r'^\d{5,6}',value,re.IGNORECASE):
        return 'группа ОВЗ'
    elif value == 'Мастер маникюра':
        return 'группа ОВЗ'

    else:
        return f'{value} неизвестная УГС'


def processing_count_ugs(data_file:str,lst_spec:str,end_folder:str):
    error_df = pd.DataFrame(columns=['Лист', 'Ошибка'])
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)

    spec_df = pd.read_excel(lst_spec,sheet_name='Выпадающие списки')

    req_wb = openpyxl.load_workbook(data_file)
    lst_sheets = req_wb.sheetnames

    dict_df = dict()
    dict_result_df = dict()

    for sheet in lst_sheets:
        print(sheet)
        temp_df = pd.read_excel(data_file, sheet_name=sheet)
        temp_df.dropna(how='all',inplace=True)

        result_df = pd.DataFrame(
            temp_df['Unnamed: 0'].values.reshape(-1, 4),
            columns=['ОГПС', 'Количество выпускников', 'Средняя зарплата', 'Процент трудоустройства']
        )

        result_df['matched'] = result_df['ОГПС'].apply(
            lambda x: find_match(x, spec_df['Код и наименование профессии, специальности'].tolist())
        )
        result_df['УГС'] = result_df['matched'].apply(extract_ugs)

        svod_df = pd.pivot_table(result_df,index=['УГС'],
                                 values=['Количество выпускников','Процент трудоустройства'],
                                 aggfunc={'Количество выпускников':'sum','Процент трудоустройства':'mean'})
        svod_df['Процент трудоустройства'] =svod_df['Процент трудоустройства'].apply(lambda x:round(x,2))
        svod_df = svod_df.reset_index()

        dict_df[f'{sheet}'] = svod_df
        dict_result_df[f'{sheet}'] = result_df


    # Объединяем
    # Добавляем колонку "Год" и объединяем
    combined = pd.concat(
        [df.assign(Год=year) for year, df in dict_df.items()],
        ignore_index=True
    )

    # Переставляем колонки в удобный порядок
    combined = combined[['Год', 'УГС', 'Количество выпускников', 'Процент трудоустройства']]
    wide = combined.pivot_table(
        index='УГС',
        columns='Год',
        values=['Количество выпускников', 'Процент трудоустройства']
    )

    wide.to_excel('data/wid.xlsx')

    combined = pd.concat(
        [df.assign(Год=year) for year, df in dict_result_df.items()],
        ignore_index=True
    )

    # Переставляем колонки в удобный порядок
    combined = combined[['Год', 'УГС', 'Количество выпускников', 'Процент трудоустройства']]

    combined.to_excel('data/comb.xlsx')



if __name__ == '__main__':
    main_data_file = 'data/огпс _ труд 2022-2025 (1).xlsx'
    main_lst_spec = 'data/Техникум №2.xlsx'
    main_end_folder = 'data/Результат'
    processing_count_ugs(main_data_file,main_lst_spec,main_end_folder)

    print('Lindy Booth')











