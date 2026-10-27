"""
Скрипт для обработки сверки лагерей
"""

import pandas as pd


def check_inn(value):
    if len(value) == 10:
        return value
    elif len(value) == 9:
        return f'0{value}'
    else:
        return f'Проверьте указанный ИНН {value}'


def prepare_inn(data_file:str, end_folder:str):
    """
    Для добавления нехватающих нулей
    :param data_file:
    :param end_folder:
    :return:
    """
    df = pd.read_excel(data_file,dtype=str)

    df['INN'] = df['INN'].astype(str)
    df['Чек_ИНН'] = df['INN'].apply(check_inn)
    df.to_excel(f'{end_folder}/Результат.xlsx',index=False)






if __name__ == '__main__':
    main_data_file = 'data/CHILDREN_RECREATION_AND_HEALTH_IMPROVEMENT (3).xlsx'
    main_end_folder = 'data/Прочее'

    prepare_inn(main_data_file, main_end_folder)
    print('Lindy Booth')
