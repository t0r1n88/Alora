"""
Скрипт для подсчета количества загрузивших данные в облако
"""

import os
import pandas as pd
from pathlib import Path
import time


def get_folder_size(folder_path):
    """
    Считает суммарный размер всех файлов в папке (рекурсивно), в байтах.
    """
    total = 0
    for dirpath, _, filenames in os.walk(folder_path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            try:
                total += os.path.getsize(fp)
            except OSError:
                pass
    return total


def build_folder_structure(root_folder, end_folder, min_size_kb=5):
    """
    Создаёт Excel с тремя листами:
      1. Все папки
      2. Папки > min_size_kb КБ
      3. Сводка: папка 1 уровня → количество больших подпапок

    Args:
        root_folder: путь к корневой папке
        output_file: путь к выходному Excel-файлу
        min_size_kb: минимальный размер подпапки в КБ
    """
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    root = Path(root_folder)
    if not root.exists():
        raise FileNotFoundError(f"Папка не найдена: {root_folder}")

    min_size_bytes = min_size_kb * 1024

    all_records = []
    big_records = []
    summary_records = []  # отдельно для сводки — включая нули

    # Все папки 1 уровня
    first_level_dirs = sorted(
        [d for d in root.iterdir() if d.is_dir()],
        key=lambda x: x.name.lower()
    )

    for first_dir in first_level_dirs:
        # Подпапки 2 уровня
        second_level_dirs = sorted(
            [d for d in first_dir.iterdir() if d.is_dir()],
            key=lambda x: x.name.lower()
        )

        # Счётчик больших подпапок для этой папки 1 уровня
        big_count = 0

        if not second_level_dirs:
            # Папка 1 уровня без подпапок → в "все" добавляем с пустым 2 уровнем
            all_records.append({
                'Папка 1 уровня': first_dir.name,
                'Папка 2 уровня': '',
                'Размер, КБ': 0
            })
        else:
            for second_dir in second_level_dirs:
                size_bytes = get_folder_size(second_dir)
                size_kb = round(size_bytes / 1024, 2)

                # В таблицу "все папки"
                all_records.append({
                    'Папка 1 уровня': first_dir.name,
                    'Папка 2 уровня': second_dir.name,
                    'Размер, КБ': size_kb
                })

                # В таблицу "папки > min_size_kb"
                if size_bytes > min_size_bytes:
                    big_records.append({
                        'Папка 1 уровня': first_dir.name,
                        'Папка 2 уровня': second_dir.name,
                        'Размер, КБ': size_kb
                    })
                    big_count += 1

        # ✅ Всегда добавляем в сводку — даже если big_count == 0
        summary_records.append({
            'Папка 1 уровня': first_dir.name,
            f'Папок > {min_size_kb} КБ': big_count
        })

    # DataFrame 1: все папки
    df_all = pd.DataFrame(all_records)

    # DataFrame 2: только большие папки
    df_big = pd.DataFrame(big_records)

    # DataFrame 3: сводка (со всеми папками 1 уровня, включая нули)
    df_big.columns = ['Муниципалитет','Организация ФИО','Размер в кб']
    df_summary = pd.DataFrame(summary_records)
    df_summary.columns = ['Муниципалитет','Количество загрузивших задание']



    # Сохраняем всё в один Excel с несколькими листами
    with pd.ExcelWriter(f'{end_folder}/Результат {current_time}.xlsx', engine='openpyxl') as writer:
        df_all.to_excel(writer, sheet_name='Все папки', index=False)
        df_big.to_excel(writer, sheet_name=f'Папки больше {min_size_kb} КБ', index=False)
        df_summary.to_excel(writer, sheet_name='Сводка', index=False)

    # Вывод статистики
    print(f"   Всего папок 1 уровня: {df_all['Папка 1 уровня'].nunique()}")
    print(f"   Всего подпапок 2 уровня: {len(df_all)}")
    print(f"   Подпапок > {min_size_kb} КБ: {len(df_big)}")
    print(f"\n📊 Сводка:")
    print(df_summary.to_string(index=False))




if __name__ == '__main__':
    main_folder = 'data/Ленинград'
    main_out_folder = 'data/Результат Ленинград'
    build_folder_structure(main_folder,main_out_folder,min_size_kb=5)

    print('Lindy Booth')
