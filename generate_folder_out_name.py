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
    Расширенная версия с копированием полной структуры папок
    """
    df = pd.read_excel(name_folder_file)

    if name_column not in df.columns:
        print(f"Ошибка: Колонка '{name_column}' не найдена")
        return

    lst_fio = df[name_column].dropna().astype(str).tolist()
    print(f"Найдено {len(lst_fio)} записей")

    source_path = Path(source_folder)
    if not source_path.exists():
        print(f"Ошибка: Папка '{source_folder}' не существует")
        return

    # Определяем корневую целевую папку
    target_root = Path(end_folder)
    target_root.mkdir(parents=True, exist_ok=True)

    stats = {
        'folders_created': 0,
        'files_copied': 0,
        'files_skipped': 0,
        'errors': 0
    }

    for fio in lst_fio:
        fio = fio.strip()

        # Полностью копируем дерево папок и файлов
        target_fio_path = target_root / fio
        target_fio_path.mkdir(exist_ok=True)
        stats['folders_created'] += 1

        # Копируем только те папки и файлы, которые соответствуют заданному шаблону
        src_tree = []
        for item in source_path.rglob('*'):
            if isinstance(item, Path) and item.match(file_pattern):
                src_tree.append(item.relative_to(source_path))  # относительный путь файла или папки

        for rel_path in src_tree:
            src_item = source_path / rel_path

            # Целевое местоположение внутри целевой папки сотрудника
            dest_item = target_fio_path / rel_path

            # Если элемент является папкой, создаем её
            if src_item.is_dir():
                dest_item.mkdir(parents=True, exist_ok=True)
            elif src_item.is_file():  # Если это файл
                if dest_item.exists() and not overwrite:
                    print(f"Файл {src_item.name} уже существует в папке {fio}, пропускаем.")
                    stats['files_skipped'] += 1
                    continue

                try:
                    shutil.copy2(src_item, dest_item)
                    stats['files_copied'] += 1
                    print(f"+ {src_item.name} -> {dest_item.parent}/{src_item.name}")
                except Exception as e:
                    print(f"Ошибка при копировании {src_item.name}: {e}")
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
