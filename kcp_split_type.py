"""
Скрипт для разделения и подсчета КЦП по УГС и типам
"""
import pandas as pd
import openpyxl
import re
import time
import numpy as np


def allocate_integer(row):
    total = row['Рекомендуемое КЦП']
    if total == 0 or pd.isna(total):
        return pd.Series([0, 0])

    # Сырые значения
    raw = np.array([row['Сырые_рабочие'], row['Сырые_спецы']])

    # Берем целые части
    ints = np.floor(raw).astype(int)
    remainder = total - ints.sum()

    # Распределяем оставшиеся места по наибольшим дробным частям
    if remainder > 0:
        # Сортируем индексы по убыванию дробной части
        sorted_indices = np.argsort(raw - ints)[::-1]
        for i in range(int(remainder)):
            ints[sorted_indices[i]] += 1

    return pd.Series([ints[0], ints[1]])


def split_kcp(data_file:str,kcp_file:str,end_folder:str):

    df = pd.read_excel(data_file)

    df_plan = pd.read_excel(kcp_file)


    df = df.rename(columns={'Названия строк': 'Строка', 'Процент': 'Значение'})

    # 1. Создаем колонку 'УГС' (заполняем только там, где есть код направления)
    # Проверяем, начинается ли строка с цифр (шаблон: 2 цифры, точка, 2 цифры, точка, 2 цифры)
    is_ugs = df['Строка'].str.contains(r'^\d{2}\.\d{2}\.\d{2}', na=False)
    df['УГС'] = np.where(is_ugs, df['Строка'], np.nan)

    # Заполняем пустые значения УГС значениями сверху (ffill - forward fill)
    df['УГС'] = df['УГС'].ffill()

    # 2. Убираем сами строки с УГС (нам нужны только строки с программами)
    df_clean = df[~is_ugs].copy()

    # 3. Переименовываем названия программ, чтобы они стали короткими заголовками колонок
    df_clean['Строка'] = df_clean['Строка'].str.replace('Программа подготовки ', '', regex=False)
    # Теперь названия будут: 'квалифицированных рабочих, служащих' и 'специалистов среднего звена'

    result = df_clean.pivot_table(
        index='УГС',
        columns='Строка',
        values='Значение',
        aggfunc='first'  # Берем первое встретившееся значение (они у нас уже уникальны)
    )

    # Опционально: переименуем колонки, чтобы они выглядели красиво
    result = result.rename(columns={
        'специалистов среднего звена': 'Программа подготовки специалистов среднего звена',
        'квалифицированных рабочих, служащих': 'Программа подготовки квалифицированных рабочих, служащих'
    })

    # Сбрасываем индекс, чтобы УГС стала обычной колонкой
    result = result.reset_index()

    df = pd.merge(df_plan, result, on='УГС', how='left')

    df.fillna(0, inplace=True)
    df['Рекомендуемое КЦП'] = df['Рекомендуемое КЦП'].apply(lambda x:round(x,0))

    # Считаем сырые значения (с копейками)
    df['Сырые_рабочие'] = df['Рекомендуемое КЦП'] * df['Программа подготовки квалифицированных рабочих, служащих']
    df['Сырые_спецы'] = df['Рекомендуемое КЦП'] * df['Программа подготовки специалистов среднего звена']

    # Применяем функцию к каждой строке
    df[['КЦП_ППКРС', 'КЦП_ППССЗ']] = df.apply(allocate_integer, axis=1)

    # Оставляем только нужные колонки
    final_df = df[['УГС', 'Рекомендуемое КЦП', 'КЦП_ППКРС', 'КЦП_ППССЗ']]
    total_row = final_df.sum(axis=0, numeric_only=True)
    total_row.name = 'Итого'  # Называем строку
    final_df = pd.concat([final_df, total_row.to_frame().T])
    final_df.loc['Итого', 'УГС'] = 'Итого'




    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    final_df.to_excel(f'{end_folder}/Раскладка по УГС и типам {current_time}.xlsx',index=False)
















if __name__ == '__main__':
    main_data_file = 'data/Список КЦП с типом.xlsx'
    main_kcp_file = 'data/расчет КЦП по УГС и типам.xlsx'
    main_end_folder = 'data/Результат'

    split_kcp(main_data_file,main_kcp_file,main_end_folder)
    print('Lindy Booth')
