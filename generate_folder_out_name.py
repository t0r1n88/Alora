"""
Скрипт для создания папок по названию из колонки таблицы и копирования туда данных
"""

import os
from pathlib import Path
import shutil
import re
import pandas as pd


# def create_structure_folder(name_folder_file:str,name_column:str,end_folder:str,name_data_file):
#     df = pd.read_excel(name_folder_file)
#     lst_fio = df['Фио'].tolist()
#
#     target = Path(end_folder) #
#     for fio in lst_fio:
#         target_fio_path = target / fio
#         target_fio_path.mkdir(exist_ok=True)
#         target_file = target_fio_path / f'Список класса.xlsx'
#         shutil.copy2(name_data_file, target_file)  # копируем файл в папку

def create_structure_folder_advanced(
        name_folder_file: str,
        name_column: str,
        end_folder: str,
        source_folder: str,
        file_pattern: str = "*",  # паттерн для фильтрации файлов (например "*.xlsx")
        recursive: bool = False,  # искать файлы в подпапках
        overwrite: bool = False  # перезаписывать существующие файлы
):
    """
    Расширенная версия с дополнительными параметрами
    """
    df = pd.read_excel(name_folder_file)

    if name_column not in df.columns:
        print(f"Ошибка: Колонка '{name_column}' не найдена")
        return

    lst_fio = df[name_column].dropna().tolist()
    print(f"Найдено {len(lst_fio)} записей")

    source_path = Path(source_folder)
    if not source_path.exists():
        print(f"Ошибка: Папка '{source_folder}' не существует")
        return

    # Поиск файлов с учетом паттерна и рекурсивности
    if recursive:
        files_to_copy = list(source_path.rglob(file_pattern))
    else:
        files_to_copy = list(source_path.glob(file_pattern))

    # Оставляем только файлы
    files_to_copy = [f for f in files_to_copy if f.is_file()]

    if not files_to_copy:
        print(f"Внимание: Нет файлов, соответствующих паттерну '{file_pattern}'")
        return

    print(f"Найдено файлов для копирования: {len(files_to_copy)}")

    target_root = Path(end_folder)
    target_root.mkdir(exist_ok=True)

    stats = {
        'folders_created': 0,
        'files_copied': 0,
        'files_skipped': 0,
        'errors': 0
    }

    for fio in lst_fio:
        fio = str(fio).strip()
        target_fio_path = target_root / fio
        target_fio_path.mkdir(exist_ok=True)
        stats['folders_created'] += 1

        for file_path in files_to_copy:
            target_file = target_fio_path / file_path.name

            # Проверяем, существует ли уже файл
            if target_file.exists() and not overwrite:
                print(f"  Файл {file_path.name} уже существует в папке {fio}, пропускаем")
                stats['files_skipped'] += 1
                continue

            try:
                shutil.copy2(file_path, target_file)
                stats['files_copied'] += 1
                print(f"  + {file_path.name} -> {fio}/")
            except Exception as e:
                print(f"  Ошибка при копировании {file_path.name}: {e}")
                stats['errors'] += 1

        print(f"✓ Папка '{fio}' обработана")

    # Выводим статистику
    print("\n" + "=" * 50)
    print("СТАТИСТИКА:")
    print(f"Создано папок: {stats['folders_created']}")
    print(f"Скопировано файлов: {stats['files_copied']}")
    print(f"Пропущено файлов: {stats['files_skipped']}")
    print(f"Ошибок: {stats['errors']}")
    print("=" * 50)


if __name__ == '__main__':

    main_name_folder_file = 'data/2026-03-10 Курс для педагогов-психологов Цифровизация профессиональной .xlsx'
    main_name_column = 'Фио'
    main_end_folder ='data/Результат'
    main_data_file = 'data/Список класса.xlsx'
    main_source_folder = 'data/Данные для копирования'
    create_structure_folder_advanced(
        name_folder_file=main_name_folder_file,
        name_column=main_name_column,
        end_folder='data/Результат',
        source_folder=main_source_folder,
        file_pattern="*",  # копировать только Excel файлы
        recursive=True,
        overwrite=True
    )
    print('Lindy Booth')
