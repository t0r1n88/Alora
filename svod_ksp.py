"""
Скрипт для создания свода по кцп
"""
import pandas as pd
import openpyxl
import re
import time

def extract_ugs(value):
    if pd.isna(value):
        return 'Отсутствует сопоставление ФГОС для данного ОКЗ'

    value = str(value).strip()
    if value.startswith('01.'):
        return '01.00.00 Математика и механика'
    elif value.startswith('02.'):
        return '02.00.00 Компьютерные и информационные науки'
    elif value.startswith('03.'):
        return '03.00.00 Физика и астрономия'
    elif value.startswith('04.'):
        return '04.00.00 Химия'
    elif value.startswith('05.'):
        return '05.00.00 Науки о Земле'
    elif value.startswith('06.'):
        return '06.00.00 Биологические науки'
    elif value.startswith('07.'):
        return '07.00.00 Архитектура'
    elif value.startswith('08.'):
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
    elif value.startswith('14.'):
        return '14.00.00 Ядерная энергетика и технологии'
    elif value.startswith('15.'):
        return '15.00.00 Машиностроение'
    elif value.startswith('16.'):
        return '16.00.00 Физико-технические науки и технологии'
    elif value.startswith('17.'):
        return '17.00.00 Оружие и системы вооружения'
    elif value.startswith('18.'):
        return '18.00.00 Химические технологии'
    elif value.startswith('19.'):
        return '19.00.00 Промышленная экология и биотехнологии'
    elif value.startswith('20.'):
        return '20.00.00 Техносферная безопасность и природообустройство'
    elif value.startswith('21.'):
        return '21.00.00 Прикладная геология, горное дело, нефтегазовое дело и геодезия'
    elif value.startswith('22.'):
        return '22.00.00 Технологии материалов'
    elif value.startswith('23.'):
        return '23.00.00 Техника и технологии наземного транспорта'
    elif value.startswith('24.'):
        return '24.00.00 Авиационная и ракетно-космическая техника'
    elif value.startswith('25.'):
        return '25.00.00 Аэронавигация и эксплуатация авиационной и ракетно-космической техники'
    elif value.startswith('26.'):
        return '26.00.00 Техника и технологии кораблестроения и водного транспорта'
    elif value.startswith('27.'):
        return '27.00.00 Управление в технических системах'
    elif value.startswith('28.'):
        return '28.00.00 Нанотехнологии и наноматериалы'
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
    elif value.startswith('41.'):
        return '41.00.00 Политические науки и регионоведение'
    elif value.startswith('42.'):
        return '42.00.00 Средства массовой информации и информационно-библиотечное дело'
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
    elif value.startswith('56.'):
        return '56.00.00 Военное управление'
    elif value.startswith('57.'):
        return '57.00.00 Обеспечение государственной безопасности'
    elif re.search(r'^\d{5,6}', value, re.IGNORECASE):
        return 'группа ОВЗ'
    elif value == 'Мастер маникюра':
        return 'группа ОВЗ'

    else:
        return f'{value} неизвестная УГС'

def filter_group(g: pd.DataFrame) -> pd.DataFrame:

    REP_COL = 'Наличие в Республике'
    LINK_COL = 'Тип связи'

    # 1. Приоритет: "готовят в Республике"
    rep = g[g[REP_COL] == 'готовят в Республике']
    if len(rep) > 0:
        # если среди них есть профильные — оставляем только профильные
        prof = rep[rep[LINK_COL] == 'Профильная работа']
        if len(prof) > 0:
            return prof
        return rep
    # 2. Иначе — профильные из всей группы
    prof = g[g[LINK_COL] == 'Профильная работа']
    if len(prof) > 0:
        return prof
    # 3. Fallback — оставляем первую строку (на случай отсутствия профильных)
    return g.head(1)

def split_sum(g: pd.DataFrame) -> pd.DataFrame:
    SUM_COL = 'Значение из колонки Сумма строки листа Замещ_2030'
    GRP_COL = 'Код начальной группы занятий'

    g = g.copy()
    n = len(g)
    if n <= 1:
        return g
    total = int(g[SUM_COL].iloc[0])  # предполагаем одинаковое значение в группе
    base, rem = divmod(total, n)
    # первые `rem` строк получают base+1, остальные base
    values = [base + 1] * rem + [base] * (n - rem)
    g[SUM_COL] = values
    return g

def create_svod_kcp(data_file: str,spec_data_file:str ,end_folder: str):

    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)
    df = pd.read_excel(data_file, dtype=str, sheet_name='Объединённая таблица')

    spec_df = pd.read_excel(spec_data_file,sheet_name='Общий список') # специальности
    spec_df = spec_df.drop_duplicates(subset=['Код и наименование'])
    spec_df[['Код', 'Название']] = spec_df['Код и наименование'].str.split(' ', n=1, expand=True)

    mask = df['Первые цифры'].str.startswith(('3', '4', '5', '6', '7', '8'), na=False)
    df = df[mask]
    df['УГС'] = df['Код профессии/специальности'].apply(extract_ugs)
    # df['Сумма строки'] = df['Сумма строки'].astype(str).str.replace('.', ',', regex=False)
    df['Сумма строки'] = df['Сумма строки'].astype(float)
    # df['Значение из колонки Сумма строки листа Замещ_2030'] = df['Значение из колонки Сумма строки листа Замещ_2030'].astype(int)
    # df['Сумма строки'] = df['Сумма строки'].astype(int)
    df['Сумма строки'] = df['Сумма строки'].apply(lambda x:round(x,0))
    df.rename(columns={'Сумма строки':'Значение из колонки Сумма строки листа Замещ_2030'},inplace=True)
    df = df.drop(columns=df.columns[3:24])

    # добавляем специальности
    result = pd.merge(df, spec_df, left_on='Код профессии/специальности', right_on='Код', how='outer',indicator=True)
    result = result[result['_merge'] != 'right_only'] # отбрасываем овз
    result.drop(columns=['База','Финансовая основа обучения','Форма обучения','План приема','КЦП','Целевые договора','План приема Профессионалитет',
                         'Всего заявлений','подано целевые','подано Профессионалитет','участники СВО','дети участников СВО','подано Госулуги',
                         'УГС_y','ОУ','Код начальной группы  ОКЗ','Наименование начальной группы ОКЗ','Код','Название','Код и наименование'],inplace=True)
    result['_merge'] = result['_merge'].replace({'left_only':'отсутствует в Республике','both':'готовят в Республике'})
    result.rename(columns={'_merge':'Наличие в Республике','УГС_x':'УГС'},inplace=True)

    base_df = result.copy() # делаем копию

    SUM_COL = 'Значение из колонки Сумма строки листа Замещ_2030'
    GRP_COL = 'Код начальной группы занятий'
    REP_COL = 'Наличие в Республике'
    LINK_COL = 'Тип связи'

    base_df = (
        result.groupby(GRP_COL, group_keys=False)
        .apply(lambda g: split_sum(filter_group(g)))
        .reset_index(drop=True)
    )

    reg_df = base_df[base_df['Наличие в Республике'] == 'готовят в Республике']
    not_reg_df = base_df[base_df['Наличие в Республике'] == 'отсутствует в Республике']


    base_svod_df = pd.pivot_table(base_df,index=['УГС'],
                                 values=['Значение из колонки Сумма строки листа Замещ_2030'],
                                 aggfunc='sum',
                                  margins=True,margins_name='Итого')


    reg_svod_df = pd.pivot_table(reg_df,index=['УГС'],
                                 values=['Значение из колонки Сумма строки листа Замещ_2030'],
                                 aggfunc='sum',
                                 margins=True,margins_name='Итого')

    not_reg_svod_df = pd.pivot_table(not_reg_df,index=['УГС'],
                                 values=['Значение из колонки Сумма строки листа Замещ_2030'],
                                 aggfunc='sum',
                                     margins=True,margins_name='Итого')

    with pd.ExcelWriter(f'{end_folder}/Итоговый свод {current_time}.xlsx') as writer:
        base_svod_df.to_excel(writer,sheet_name='Свод по всем',index=True)
        reg_svod_df.to_excel(writer,sheet_name='Свод по местным проф',index=True)
        not_reg_svod_df.to_excel(writer,sheet_name='Свод по отсутствующим проф',index=True)
        base_df.to_excel(writer,sheet_name='список для свода',index=False)
        result.to_excel(writer,sheet_name='Сырые данные',index=False)



if __name__ == '__main__':
    main_data_file = 'data/Соединение ОКЗ из листа Замещ_2030 и ОКЗ из Переходника Минруда.xlsx'
    main_spec_data_file = 'data/Итоговый файл 28_09_2026.xlsx'
    main_end_folder = 'data/Результат'

    create_svod_kcp(main_data_file,main_spec_data_file, main_end_folder)

    print('Lindy Booth')
