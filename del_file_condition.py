"""
Удаление файлов по условию
"""
import os







def processing_delete_files_conditions(data_folder:str, str_conditions :str):

    for dirpath, dirnames, filenames in os.walk(data_folder):
        for file in filenames:
            if str_conditions in file:
                if os.path.exists(f'{dirpath}/{file}'):
                    os.remove(f'{dirpath}/{file}')










if __name__ == '__main__':
    main_data_folder = 'c:/Users/1/Downloads/Telegram Desktop/ChatExport_2026-03-03/photos/'
    main_data_folder = 'c:/Users/1/Downloads/Telegram Desktop/ChatExport_2026-03-03/video_files/'
    main_str_conditions = 'thumb'
    processing_delete_files_conditions(main_data_folder,main_str_conditions)
    print('Lindy Booth')
