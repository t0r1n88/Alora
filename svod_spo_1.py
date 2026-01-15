"""
Скрипт для подсчета данных
"""
import pandas as pd
import os


def extract_data_spo_one(data_folder:str,end_folder:str):
    dct_df = {
        'Очное ССЗ':pd.DataFrame(columns=['1','2','3''4','5','6''7','8','9','10','11','12']),
        'Очное КРС':pd.DataFrame(columns=['1','2','3''4','5','6''7','8','9','10','11','12']),
        'Очно-Заочно ССЗ':pd.DataFrame(columns=['1','2','3''4','5','6''7','8','9','10','11','12']),
        'Очно-Заочно КРС':pd.DataFrame(columns=['1','2','3''4','5','6''7','8','9','10','11','12']),
        'Заочное ССЗ':pd.DataFrame(columns=['1','2','3''4','5','6''7','8','9','10','11','12']),
        'Заочное КРС':pd.DataFrame(columns=['1','2','3''4','5','6''7','8','9','10','11','12']),
    }







    for dirpath, dirnames, filenames in os.walk(data_folder):
        for file in filenames:
            if not file.startswith('~$') and (file.endswith('.xlsx')):
                name_file = file.split('.xlsx')[0].strip()
                print(name_file)
                for name_sheet,base_df in dct_df.items():
                    # Перебираем листы






if __name__ == '__main__':
    main_data_folder = 'data/СПО'
    main_end_folder = 'data/Результат'
    extract_data_spo_one(main_data_folder,main_end_folder)
