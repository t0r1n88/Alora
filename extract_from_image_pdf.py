import re
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance
import os
import numpy as np
import pandas as pd

# Укажите путь к папке bin Poppler (для Windows)
POPPLER_PATH = r"c:/poppler-25.12.0/Library/bin/"  # ИЗМЕНИТЕ НА ВАШ ПУТЬ!

# Укажите путь к Tesseract OCR (для Windows)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def enhance_image(image):
    """
    Улучшает качество изображения для OCR
    """
    # Конвертируем в оттенки серого
    if image.mode != 'L':
        gray = image.convert('L')
    else:
        gray = image

    # Увеличиваем контрастность
    enhancer = ImageEnhance.Contrast(gray)
    enhanced = enhancer.enhance(2.0)

    # Увеличиваем резкость
    enhancer = ImageEnhance.Sharpness(enhanced)
    enhanced = enhancer.enhance(1.5)

    return enhanced




def extract_info_from_pdf(pdf_path, use_opencv=False):
    """
    Извлекает ФИО и Unti-ID из PDF с помощью OCR
    """
    info = {
        'filename': os.path.basename(pdf_path),
        'fio': None,
        'student_id': None
    }


    # Конвертируем PDF в изображения
    print("Конвертация PDF в изображения...")
    try:
        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH, dpi=300)
    except Exception as e:
        print(f"Ошибка при конвертации PDF: {e}")
        print("Проверьте путь к Poppler!")
        return info

    # Обрабатываем каждую страницу
    full_text = ""
    for i, image in enumerate(images):
        print(f"Обработка страницы {i + 1} из {len(images)}...")

        # Улучшаем качество изображения
        enhanced_image = enhance_image(image)

        # Выполняем OCR
        try:
            # Настройки для русского языка
            custom_config = r'--oem 3 --psm 6 -l rus+eng'
            page_text = pytesseract.image_to_string(enhanced_image, config=custom_config)

            # Добавляем текст в общую строку
            full_text += page_text + "\n"

            # # Выводим распознанный текст для отладки (первые 200 символов)
            # print(f"  Распознано {len(page_text)} символов")
            # if i == 0:  # Только для первой страницы
            #     print(f"  Первые 200 символов: {page_text[:200]}")

        except Exception as e:
            print(f"Ошибка OCR на странице {i + 1}: {e}")

    # Поиск ФИО в тексте
    print("\nПоиск ФИО и ID в распознанном тексте...")

    # Паттерны для поиска ФИО
    fio_patterns = [r'ФИО.*?:\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)'
    ]

    for pattern in fio_patterns:
        fio_match = re.search(pattern, full_text, re.IGNORECASE | re.MULTILINE)
        if fio_match:
            # Берем последнюю группу (само ФИО)
            info['fio'] = fio_match.group(fio_match.lastindex or 1)
            print(f"  Найдено ФИО: {info['fio']}")
            break

    # Поиск Unti-ID
    id_patterns = [
        r'Unti-ID\)?:\s*(\d+)',
            r'Unti[-\s]*ID[:\s]*(\d+)'  # Unti-ID: 123456
    ]

    for pattern in id_patterns:
        id_match = re.search(pattern, full_text, re.IGNORECASE)
        if id_match:
            info['student_id'] = id_match.group(1)
            print(f"  Найден ID: {info['student_id']}")
            break



    return info


def save_debug_image(image, page_num):
    """Сохраняет изображение для отладки"""
    debug_dir = "debug_images"
    if not os.path.exists(debug_dir):
        os.makedirs(debug_dir)

    image.save(os.path.join(debug_dir, f"page_{page_num}.png"))
    print(f"  Изображение сохранено: debug_images/page_{page_num}.png")


# Основная функция
def main():
    # Путь к вашему PDF файлу
    data_folder = "data/Анкеты ЧТОТиБ/"  # ИЗМЕНИТЕ НА ВАШ ПУТЬ

    base_df = pd.DataFrame(columns=[['№','UNIT ID','ФИО','Статус','Комментарий Провайдера','Имя файла','Название площадки','ФИО Менеджера','Дата аттестация']])


    for dirpath, dirnames, filenames in os.walk(data_folder):
        for file in filenames:
            if file.endswith('.pdf'):
                name_file = file.split('.pdf')
                print(f"Обработка файла: {name_file}")
                print("-" * 50)

                # Извлекаем информацию
                result = extract_info_from_pdf(f'{dirpath}/{file}', use_opencv=False)

                # Выводим результаты
                print("\n" + "=" * 60)
                print("РЕЗУЛЬТАТЫ ИЗВЛЕЧЕНИЯ:")
                print("=" * 60)

                print(f"Имя файла: {result['filename']}")

                print(f"\nФИО (из OCR): {result.get('fio', 'НЕ НАЙДЕНО')}")
                print(f"Unti-ID (из OCR): {result.get('student_id', 'НЕ НАЙДЕНО')}")

                # Итоговый результат
                print("\n" + "=" * 60)
                print("ИТОГОВЫЙ РЕЗУЛЬТАТ:")
                print("=" * 60)
                print(f"ФИО обучающегося: {result.get('fio', 'Не удалось распознать')}")
                print(f"Unti-ID: {result.get('student_id', 'Не удалось распознать')}")




if __name__ == "__main__":
    main()