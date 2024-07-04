"""
Скрипт для обработки данных мониторинга уровня использования Сферум по районам
"""
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import time
import re

class NoNameColumn(BaseException):
    """
    Класс для обозначения ситуации когда не найдена колонка с нужным именем
    """
    pass

"""
Алгоритм
1) Отфильтровать данные по республике
2) Создать документ openpyxl
3) Отфильтровать данные по району и создать лист с соответствующим названием
4) Поместить данные на этот лист и закрасить красным строки где в определенных колонках значение меньше целевого показателя

Примечания
Размер целевого показателя должен вводиться через поле ввода  на данный момент 32,63%
Названия листа с которого будут забираться данные тоже должны вводится
"""
def process_split_highlighting_threshold(path:str,name_sheet:str,region:str,path_to_end_folder:str)-> None:
    df = pd.read_excel(path,dtype={'ИНН':str})
    lst_need_columns = ['id','Субъект РФ','Наименование АТО','Краткое наименование организации',
                        'Полное наименование организации','Адрес организации','ИНН','Тип поселения',
                        'Статус организации\n(юр.лицо/филиал)','Статус\nфункционирования',
                        'Государственная/негосударственная организация (ГОУ/НОУ)',
                        'Численность обучающихся\nвсего, чел. (Раздел 1.3. стр.01 гр.3)',
                        'Численность педработников - всего, чел. (Раздел 3.1 стр.06 гр.3)',
                        'Обучающиеся Сферум (чел)','Обучающиеся МШ (чел)','Обучающиеся Cферум или МШ (чел)',
                        'Обучающиеся, использовавшие за последние 12 месяцев Сферум',
                        ' Педагоги, использовавшие за последние 12 месяцев Сферум','Обучающиеся, использовавшие за последние 12 месяцев МШ',
                        ' Педагоги, использовавшие за последние 12 месяцев МШ','Педагоги, использовавшие за последние 12 месяцев Сферум или МШ (чел)',
                        'Обучающиеся, использовавшие за последние 12 месяцев Сферум % от общего числа',
                        'Педагоги, использовавшие за последние 12 месяцев Сферум % от общего числа',
                        'Обучающиеся, использовавшие за последние 12 месяцев МШ % от общего числа',
                        'Педагоги, использовавшие за последние 12 месяцев МШ % от общего числа']

    diff_set = set(lst_need_columns).difference(set(df.columns))
    if len(diff_set) > 0 :
        raise NoNameColumn

    # Отбираем нужные колонки
    df = df[lst_need_columns]
    df = df[df['Субъект РФ'] == region] # Фильтруем регион
    lst_unqiue_ruo = df['Наименование АТО'].unique() # получаем список РУО

    used_name_sheet = set()  # множество для хранения значений которые уже были использованы
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    # Создаем документ и листы
    wb = openpyxl.Workbook()
    for idx,name_ruo in enumerate(lst_unqiue_ruo):
        temp_df = df[df['Наименование АТО'] == name_ruo]
        short_value = name_ruo[:20]  # получаем обрезанное значение
        short_value = re.sub(r'[\[\]\'+()<> :"?*|\\/]', '_', short_value)

        if short_value in used_name_sheet:
            short_value = f'{short_value}_{idx}'  # добавляем окончание
        wb.create_sheet(short_value, index=idx)  # создаем лист
        used_name_sheet.add(short_value)
        for row in dataframe_to_rows(temp_df, index=False, header=True):
            wb[short_value].append(row)

            # Устанавливаем автоширину для каждой колонки
            for column in wb[short_value].columns:
                max_length = 0
                column_name = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                wb[short_value].column_dimensions[column_name].width = adjusted_width
        wb.save(f'{path_to_end_folder}/Районы МШ и Сферум {current_time}.xlsx')
        wb.close()


if __name__ == '__main__':
    main_df = 'data/Расчет_2.1_2.2_2.3_июнь_2024 г.xlsx'
    main_name_sheet = 'Процент пользователей'
    main_region = 'Республика Бурятия'
    main_end_folder = 'data/Результат'

    process_split_highlighting_threshold(main_df, main_name_sheet, main_region,main_end_folder)

    print('Lindy Booth')


