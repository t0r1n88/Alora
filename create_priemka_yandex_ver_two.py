"""
Обработка Яндекс таблицы новой версии
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
    elif re.search(r'\bОчное\b',form_str):
        return 'очная'
    elif re.search(r'\очно[- ]заочн|очн[оы]е?[- ]заочн|оч[-.]заоч',form_str,re.IGNORECASE):
        return 'очно-заочная'
    elif re.search(r'\bзаочная\b',form_str,re.IGNORECASE):
        return 'заочная'
    elif re.search(r'\bзаочное\b',form_str,re.IGNORECASE):
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
    elif re.search(r'коммер|платн|оплат|договор|ком|полн',form_str,re.IGNORECASE):
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
    elif re.search(r'коррек',form_str,re.IGNORECASE):
        return 'ОВЗ'

    else:
        return f'{value} неизвестный уровень образования'


def extract_ugs(value):
    if pd.isna(value):
        return 'Не заполнен Код и наименование'

    value = str(value).strip()

    if value.startswith('08.'):
        return '08.00.00 Техника и технологии строительства'
    elif value.startswith('09.'):
        return '09.00.00 Информатика и вычислительная техника'
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

    else:
        return f'{value} неизвестная УГС'






def generate_data_for_priem_yandex(data_file:str,end_folder:str):
    """

    :param data_file: Файл из яндекс диска
    :param end_folder: Конечная папка
    """
    error_df = pd.DataFrame(columns=['Лист', 'Ошибка'])
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)

    req_wb = openpyxl.load_workbook(data_file)
    lst_sheets = req_wb.sheetnames
    lst_cols = ['ОУ', 'Код и наименование', 'База',
                'Финансовая основа обучения', 'Форма обучения', 'План приема',
                'КЦП', 'Целевые договора', 'План приема Профессионалитет',
                'Всего заявлений', 'подано целевые', 'подано Профессионалитет',
                'участники СВО', 'дети участников СВО','подано Госулуги']
    main_df = pd.DataFrame(columns=lst_cols)
    for sheet in lst_sheets:
        print(sheet)
        temp_df = pd.read_excel(data_file, sheet_name=sheet, skiprows=1)
        temp_df = temp_df.dropna(axis=1, how='all')



        # добавляем колонку ОУ если ее нет
        lst_ou = [col for col in temp_df.columns if 'филиалы' in str(col).lower()]
        if len(lst_ou) == 0:
            temp_df.insert(0,'Наименование ОУ, филиалы','1')
        else:
            temp_df.rename(columns={lst_ou[0]:'Наименование ОУ, филиалы'},inplace=True)


        # удаляем колонки относящиеся к Зачислено
        if 'Зачислено (принято обучающихся) ' in temp_df.columns:
            idx = temp_df.columns.get_loc('Зачислено (принято обучающихся) ')
            temp_df = temp_df.drop(columns=temp_df.columns[idx:idx + 5])

        # удаляем колонки относящиеся к Зачислено
        if 'Зачислено (принято обучающихся) с сентября по октябрь' in temp_df.columns:
            idx = temp_df.columns.get_loc('Зачислено (принято обучающихся) с сентября по октябрь')
            temp_df = temp_df.drop(columns=temp_df.columns[idx:idx + 5])


        # Находим последнюю колонку Всего
        indices = [i for i, col in enumerate(temp_df.columns)
                   if temp_df[col].astype(str).str.contains('Всего', na=False).any()]
        if len(indices) == 0:
            temp_error_df = pd.DataFrame(columns=['Лист', 'Ошибка'], data=[[sheet, 'Отсутствует строка со значением Всего']])
            error_df = pd.concat([error_df, temp_error_df])
            continue
        last_idx = indices[-1] # берем последнее всего

        #
        # print(f"Индексы колонок со словом 'Всего': {indices}")
        # temp_df.to_excel('data/gg.xlsx',index=False)
        # print(temp_df[temp_df.columns[last_ind]])



        # print('Проверка last_idx')
        # print(last_idx)
        # print(temp_df.shape)
        # #
        # print(temp_df[temp_df.columns[last_idx]])
        # print('Окончание проверки last_idx')
        # temp_df.to_excel('data/res.xlsx',index=False)
        # raise ZeroDivisionError

        if not last_idx:
            temp_error_df = pd.DataFrame(columns=['Лист', 'Ошибка'], data=[[sheet, 'В конце не найдена колонка Всего']])
            error_df = pd.concat([error_df, temp_error_df])
            continue




        temp_df = pd.concat([
            temp_df.iloc[:, :9],  # первые 9 колонок
            temp_df.iloc[:, last_idx:]  # последние 6 колонок
        ], axis=1)

        temp_df.dropna(inplace=True, thresh=9)
        temp_df = temp_df[temp_df['Наименование ОУ, филиалы'].notna()]  # отбрасываем строки у которых не записан ОУ
        temp_df = temp_df[temp_df['КОД и наименование профессий и специальностей'].notna()]  # отбрасываем строки у которых не записан ОУ

        if len(temp_df) == 0:
            temp_error_df = pd.DataFrame(columns=['Лист', 'Ошибка'], data=[[sheet, 'На заполнен лист']])
            error_df = pd.concat([error_df, temp_error_df])
            continue
        pattern = '|'.join(['итог', 'всего'])
        temp_df = temp_df[~temp_df['Наименование ОУ, филиалы'].str.contains(pattern, case=False, na=False)]
        temp_df = temp_df[~temp_df['КОД и наименование профессий и специальностей'].str.contains(pattern, case=False, na=False)]
        temp_df = temp_df.applymap(
            lambda x: re.sub(r'\s+', ' ', x) if isinstance(x, str) else x)  # очищаем от лишних пробелов
        temp_df = temp_df.applymap(
            lambda x: x.strip() if isinstance(x, str) else x)  # очищаем от пробелов в начале и конце

        check_lst = [col for col in temp_df.columns if 'статусом' in str(col)]
        if len(check_lst)> 1:
            temp_error_df = pd.DataFrame(columns=['Лист', 'Ошибка'], data=[[sheet, 'Несколько колонок Из них со статусом Новое']])
            error_df = pd.concat([error_df, temp_error_df])
            continue
        elif len(check_lst) == 1:
            temp_df.drop(columns=check_lst[0],inplace=True)



        # Проверяем наличие колонки с данными госуслуг
        gos_check_lst = [col for col in temp_df.columns if 'Госуслуги' in str(col)]
        if len(gos_check_lst)> 1:
            temp_error_df = pd.DataFrame(columns=['Лист', 'Ошибка'], data=[[sheet, 'Несколько колонок Количество поданных через Госусуги']])
            error_df = pd.concat([error_df, temp_error_df])
            continue
        elif len(gos_check_lst) == 0:
            temp_df['Госуслуги'] = 0

        # особое исключение для БРИТ
        if sheet in ('БРИТ','КТИРНЗ'):
            temp_df.drop(columns=['Госуслуги'],inplace=True)
            temp_df.columns = ['ОУ', 'Код и наименование', 'База',
                'Финансовая основа обучения', 'Форма обучения', 'План приема',
                'КЦП', 'Целевые договора', 'План приема Профессионалитет',
                'Всего заявлений','подано Госулуги', 'подано целевые', 'подано Профессионалитет',
                'участники СВО', 'дети участников СВО']

        else:

            temp_df.columns = lst_cols
        temp_df['ОУ'] = sheet

        temp_df = temp_df.reindex(columns=['ОУ','Код и наименование','База',
                    'Финансовая основа обучения','Форма обучения','План приема',
                    'КЦП','Целевые договора','План приема Профессионалитет',
                    'Всего заявлений','подано Госулуги','подано целевые',
                    'подано Профессионалитет','участники СВО','дети участников СВО'])

        main_df = pd.concat([main_df, temp_df])

    main_df['Форма обучения'] = main_df['Форма обучения'].apply(extract_educ_form)
    main_df['Финансовая основа обучения'] = main_df['Финансовая основа обучения'].apply(extract_pay_form)
    main_df['База'] = main_df['База'].apply(extract_level_educ)
    main_df.sort_values(by='ОУ', inplace=True)
    main_df.fillna(0, inplace=True)
    main_df['УГС'] = main_df['Код и наименование'].apply(extract_ugs)
    main_df[main_df.columns[-10:-1]] = main_df[main_df.columns[-10:-1]].apply(pd.to_numeric, errors='coerce').fillna(
        0).astype(int)

    svod_df = pd.pivot_table(main_df,
                             values=['Всего заявлений', 'подано Госулуги', 'подано целевые', 'подано Профессионалитет',
                                     'участники СВО', 'дети участников СВО'],
                             index=['ОУ'],
                             aggfunc='sum',
                             fill_value=0, )

    total_row = svod_df.sum(axis=0, numeric_only=True)
    total_row.name = 'Итого'  # Называем строку
    svod_df = pd.concat([svod_df, total_row.to_frame().T])
    svod_df = svod_df.reindex(
        columns=['Всего заявлений', 'подано Госулуги', 'подано Профессионалитет', 'подано целевые', 'участники СВО',
                 'дети участников СВО'])

    with pd.ExcelWriter(f'{end_folder}/Итоговый файл {current_date}.xlsx') as writer:
        svod_df.to_excel(writer, sheet_name='Свод', index=True)
        main_df.to_excel(writer, sheet_name='Общий список', index=False)

    print(error_df)
    error_df.to_excel(f'{end_folder}/Ошибки_{current_time}.xlsx', index=False)




if __name__ == '__main__':
    main_data_file = 'data/ПРИЕМНАЯ КАМПАНИЯ 2026-2027.xlsx'
    main_end_result_folder = 'data/Результат'

    generate_data_for_priem_yandex(main_data_file,main_end_result_folder)

    print('Lindy Booth')












