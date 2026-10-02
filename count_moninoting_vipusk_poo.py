"""
Скрипт для подсчета данных по трудоустройству по ПОО
"""

import pandas as pd
import openpyxl
import time
import re
import numpy as np


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

    return value

def extract_ugs(value):
    if pd.isna(value):
        return 'Не заполнен Код и наименование'

    value = str(value).strip()

    if value.startswith('08.'):
        return '08.00.00 Техника и технологии строительства'
    elif value.startswith('09.'):
        return '09.00.00 Информатика и вычислительная техника'
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


def convert_to_int(value):
    try:
        return int(re.sub(r'\D', '', value))


    except:
        return 0

def convert_to_int_percent(value):

    try:
        return float(value) * 100
    except:
        return 0



def parse_with_positions(raw_data):
    raw_data = [str(x).strip() for x in raw_data if str(x).strip()]

    # Находим позиции всех техникумов
    tech_positions = [(i, x) for i, x in enumerate(raw_data) if x.startswith('*')]

    records = []

    for idx, (pos, tech) in enumerate(tech_positions):
        # Конец блока — следующая позиция техникума или конец списка
        end = tech_positions[idx + 1][0] if idx + 1 < len(tech_positions) else len(raw_data)

        # Данные после техникума до следующего техникума
        block = raw_data[pos + 1:end]

        # Разбиваем блок на порции по 4 значения
        for i in range(0, len(block), 4):
            chunk = block[i:i + 4]
            if len(chunk) == 4:
                records.append({
                    'ПОО': tech,
                    'Название': chunk[0],
                    'Количество выпускников': convert_to_int(chunk[1]),
                    'Средняя зарплата': convert_to_int(chunk[2]),
                    'Процент трудоустройства': convert_to_int_percent(chunk[3]),
                })

    return pd.DataFrame(records)



def processing_poo_ugs(data_file:str, lst_spec:str, end_folder:str):
    error_df = pd.DataFrame(columns=['Лист', 'Ошибка'])
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)

    spec_df = pd.read_excel(lst_spec,sheet_name='Выпадающие списки')

    req_wb = openpyxl.load_workbook(data_file)
    lst_sheets = req_wb.sheetnames


    replace_dct = {'Байкальский базовый медицинский колледж Минздрава Республики Бурятия':'ББМК',
                   'Байкальский колледж недропользования':'БКН',
                   'Байкальский колледж технологий':'БКТ',
                   'Байкальский колледж туризма и сервиса':'БКТИС',
                   'Байкальский многопрофильный колледж':'БМК',
                   'Бурятская государственная сельскохозяйственная академия им. В.Р. Филиппова':'АТ БГСХА',
                   'Бурятский аграрный колледж им. М.Н. Ербанова':'БАК',
                   'Бурятский государственный университет':'Колледж БГУ',
                   'Бурятский колледж технологий и лесопользования':'БКТИЛ',
                   'Бурятский республиканский индустриальный техникум':'БРИТ',
                   'Бурятский республиканский информационно-экономический техникум':'БРИЭТ',
                   'Бурятский республиканский информационно-экономический техникум (Тункинский филиал)':'БРИЭТ',
                   'Бурятский республиканский многопрофильный техникум инновационных технологий':'БРМТИТ',
                   'Бурятский республиканский педагогический колледж':'БРПК',
                   'Бурятский республиканский техникум автомобильного транспорта':'БРТАТ',
                   'Бурятский республиканский техникум строительных и промышленных технологий':'БРТСиПТ',
                   'Бурятский республиканский хореографический колледж им. Л.П. Сахьяновой и П.Т. Абашеевой':'БРХК',
                   'Бурятский финансово-кредитный колледж':'БФКК',
                   'Восточно-Сибирский государственный университет технологий и управления':'ТК ВСГУТУ',
                   'Гусиноозерский энергетический техникум (Республика Бурятия)':'ГЭТ',
                   'Джидинский многопрофильный техникум':'ДМТ',
                   'Закаменский агропромышленный техникум':'ЗАПТ',
                   'Колледж искусств им. П.И. Чайковского (Республика Бурятия)':'КИ',
                   'Политехнический техникум (Республика Бурятия)':'ПТ',
                   'Республиканский базовый медицинский колледж им. Э.Р. Раднаева':'РБМК',
                   'Республиканский межотраслевой техникум':'РМТ',
                   'Республиканский многоуровневый колледж':'РМК',
                   'Сибирский государственный университет телекоммуникаций и информатики (Бурятский институт инфокоммуникаций)':'БИИК СибГУТИ',
                   'Техникум строительства и городского хозяйства (г. Улан-Удэ)':'ТСиГХ',
                   'Улан-Удэнский авиационный техникум':'УУАТ',
                   'Улан-Удэнский колледж железнодорожного транспорта (филиал Иркутского государственного университета путей сообщения)':'УУКЖТ',
                   'Улан-Удэнский техникум экономики, торговли и права Бурятского республиканского союза потребительских обществ':'УУТЭТиП',
                   'Колледж традиционных искусств и ремесел народов Забайкалья':'КТИРНЗ',
                   }

    for sheet in lst_sheets:
        print(sheet)
        temp_df = pd.read_excel(data_file, sheet_name=sheet)
        temp_df.dropna(how='all',inplace=True)

        raw_data = temp_df.iloc[:, 0].dropna().astype(str).tolist()

        temp_df = parse_with_positions(raw_data)
        temp_df['Специальность'] = temp_df['Название'].apply(
            lambda x: find_match(x, spec_df['Код и наименование профессии, специальности'].tolist())
        )
        temp_df['УГС'] = temp_df['Специальность'].apply(extract_ugs)
        temp_df.drop(columns=['Название'],inplace=True)

        temp_df['ПОО'] = temp_df['ПОО'].apply(lambda x:str(x).replace('*',''))
        temp_df['ПОО'] = temp_df['ПОО'].replace(replace_dct)





        svod_df = pd.pivot_table(temp_df,index=['ПОО'],
                                 values=['Количество выпускников','Процент трудоустройства'],
                                 aggfunc={'Количество выпускников':'sum','Процент трудоустройства':'mean'})
        svod_df['Процент трудоустройства'] =svod_df['Процент трудоустройства'].apply(lambda x:round(x,2))
        svod_df = svod_df.reset_index()

        total_row = svod_df.sum(axis=0, numeric_only=True)
        total_row.name = 'Итого'  # Называем строку
        svod_df = pd.concat([svod_df, total_row.to_frame().T])
        svod_df.loc['Итого','ПОО'] = 'Итого'
        svod_df.loc['Итого','Процент трудоустройства'] = ''
        dict_df = {f'Свод {sheet}':svod_df}

        temp_df = temp_df.reindex(
            columns=['ПОО', 'Специальность', 'Количество выпускников', 'Средняя зарплата',
                     'Процент трудоустройства', 'УГС'])


        lst_poo = sorted(temp_df['ПОО'].unique())
        for poo in lst_poo:
            poo_df = temp_df[temp_df['ПОО'] == poo]
            dict_df[poo] = poo_df


        with pd.ExcelWriter(f'{end_folder}/Данные {sheet}.xlsx') as writer:
            for name_sheet, df in dict_df.items():
                df.to_excel(writer,index=False,sheet_name=name_sheet)





if __name__ == '__main__':
    main_data_file = 'data/статистика трудоустройства выпускников 24-25.xlsx'
    main_lst_spec = 'data/Техникум №2.xlsx'
    main_end_folder = 'data/Результат'
    processing_poo_ugs(main_data_file, main_lst_spec, main_end_folder)

    print('Lindy Booth')











