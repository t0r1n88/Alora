from  first_prof import processing_data_first_prof # для обработки файла первой профессии
from create_svod_first_prof import generate_svod_first_prof # для сводки по приступившим к обучению
from bvb_events_rmg import create_svod_bvb # для обработки данных билета в будущее
from alora_diff_tables import find_diffrence # функция для нахождения разницы между двумя таблицами
import tkinter
import sys
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time
# pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

"""
Служебные функции в том числе для работы графического интерфейса
"""

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def make_textmenu(root):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # эта штука делает меню
    global the_menu
    the_menu = Menu(root, tearoff=0)
    the_menu.add_command(label="Вырезать")
    the_menu.add_command(label="Копировать")
    the_menu.add_command(label="Вставить")
    the_menu.add_separator()
    the_menu.add_command(label="Выбрать все")


def callback_select_all(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # select text after 50ms
    window.after(50, lambda: event.widget.select_range(0, 'end'))


def show_textmenu(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    e_widget = event.widget
    the_menu.entryconfigure("Вырезать", command=lambda: e_widget.event_generate("<<Cut>>"))
    the_menu.entryconfigure("Копировать", command=lambda: e_widget.event_generate("<<Copy>>"))
    the_menu.entryconfigure("Вставить", command=lambda: e_widget.event_generate("<<Paste>>"))
    the_menu.entryconfigure("Выбрать все", command=lambda: e_widget.select_range(0, 'end'))
    the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)


def on_scroll(*args):
    canvas.yview(*args)

def set_window_size(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Устанавливаем размер окна в 80% от ширины и высоты экрана
    if screen_width >= 3840:
        width = int(screen_width * 0.2)
    elif screen_width >= 2560:
        width = int(screen_width * 0.31)
    elif screen_width >= 1920:
        width = int(screen_width * 0.41)
    elif screen_width >= 1600:
        width = int(screen_width * 0.5)
    elif screen_width >= 1280:
        width = int(screen_width * 0.62)
    elif screen_width >= 1024:
        width = int(screen_width * 0.77)
    else:
        width = int(screen_width * 1)

    height = int(screen_height * 0.8)

    # Рассчитываем координаты для центрирования окна
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размер и положение окна
    window.geometry(f"{width}x{height}+{x}+{y}")





def select_template_folder_data():
    """
    Функция для выбора папки c данными
    :return:
    """
    global path_template_folder_data
    path_template_folder_data = filedialog.askdirectory()

def select_template_end_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_template_end_folder
    path_to_template_end_folder = filedialog.askdirectory()

def select_template_file_docx():
    """
    Функция для выбора файла Word
    :return: Путь к файлу шаблона
    """
    global file_template_docx
    file_template_docx = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def select_singe_file_template_data_xlsx():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global singe_file_template_data_xlsx
    # Получаем путь к файлу
    singe_file_template_data_xlsx = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_many_files_template_data_xlsx():
    """
    Функция для выбора нескоьких файлов с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global many_files_template_data_xlsx
    # Получаем список с файлами
    many_files_template_data_xlsx = filedialog.askopenfilenames(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))




def select_data_yandex_first_prof():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global data_yandex_first_prof
    # Получаем путь к файлу
    data_yandex_first_prof = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_result_yandex_first_prof_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_first_prof_end_folder
    path_to_first_prof_end_folder = filedialog.askdirectory()



def select_data_itog_list_first_prof():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global data_itog_list
    # Получаем путь к файлу
    data_itog_list = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_data_est_first_prof():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global data_est_first_prof
    # Получаем путь к файлу
    data_est_first_prof = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_data_moodle_first_prof():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global data_moodle_first_prof
    # Получаем путь к файлу
    data_moodle_first_prof = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))




def select_result_svod_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_svod_first_prof
    path_to_end_svod_first_prof = filedialog.askdirectory()


def processing_preparation_yandex_first_prof():
    """
    Функция для генерации документов
    """
    try:
        processing_data_first_prof(data_yandex_first_prof,path_to_first_prof_end_folder)

    except NameError:
        messagebox.showerror('',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')


def processing_create_svod_first_prof():
    """
    Функция для генерации документов
    """
    try:
        generate_svod_first_prof(data_itog_list,data_est_first_prof,data_moodle_first_prof,path_to_end_svod_first_prof)

    except NameError:
        messagebox.showerror('',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')



"""
Функции для обработки данных билета в будущее
"""

def select_data_rmg():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global data_rmg
    # Получаем путь к файлу
    data_rmg = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_data_bvb():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global data_bvb
    # Получаем путь к файлу
    data_bvb = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_bvb():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_bvb
    path_to_end_bvb = filedialog.askdirectory()

def processing_create_svod_bvb():
    """
    Функция для генерации документов
    """
    try:
        create_svod_bvb(data_bvb,data_rmg,path_to_end_bvb)

    except NameError:
        messagebox.showerror('',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')



"""
Нахождения разницы 2 таблиц
Функции  получения параметров для find_diffrenece 
"""


def select_first_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_first_diffrence
    # Получаем путь к файлу
    data_first_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'),('Excel files', '*.xlsm'), ('all files', '*.*')))


def select_second_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_second_diffrence
    # Получаем путь к файлу
    data_second_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'),('Excel files', '*.xlsm'), ('all files', '*.*')))


def select_end_folder_diffrence():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_diffrence
    path_to_end_folder_diffrence = filedialog.askdirectory()


def processing_diffrence():
    """
    Функция для получения названий листов и путей к файлам которые нужно сравнить
    :return:
    """
    # названия листов в таблицах
    try:
        first_sheet = entry_first_sheet_name_diffrence.get()
        second_sheet = entry_second_sheet_name_diffrence.get()
        # находим разницу
        find_diffrence(first_sheet, second_sheet, data_first_diffrence, data_second_diffrence,
                       path_to_end_folder_diffrence)
    except NameError:
        messagebox.showerror('Алора',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')







if __name__ == '__main__':
    window = Tk()
    window.title('Алора ver 1.0')
    # Устанавливаем размер и положение окна
    set_window_size(window)
    # window.geometry('774x760')
    # window.geometry('980x910+700+100')
    window.resizable(True, True)
    # Добавляем контекстное меню в поля ввода
    make_textmenu(window)

    # Создаем вертикальный скроллбар
    scrollbar = Scrollbar(window, orient="vertical")

    # Создаем холст
    canvas = Canvas(window, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)

    # Привязываем скроллбар к холсту
    scrollbar.config(command=canvas.yview)

    # Создаем ноутбук (вкладки)
    tab_control = ttk.Notebook(canvas)

    """
    Создаем вкладку для обработки данных билета в будущее
    """

    tab_bvb = ttk.Frame(tab_control)
    tab_control.add(tab_bvb, text='Создание сводов Билет в будущее')

    bvb_frame_description = LabelFrame(tab_bvb)
    bvb_frame_description.pack()

    lbl_hello_bvb = Label(bvb_frame_description,
                          text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                               'Для работы требуются 2 файла скачанных с сайта Билета в будущее:\n'
                               '1) Данные по Россия мои Горизонты\n'
                               '2) Сводный отчет по ученикам', width=60)
    lbl_hello_bvb.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_bvb = resource_path('logo.png')
    img_bvb = PhotoImage(file=path_to_img_bvb)
    Label(bvb_frame_description,
          image=img_bvb, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_bvb = LabelFrame(tab_bvb, text='Подготовка')
    frame_data_bvb.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать файл

    btn_bvb_rmg = Button(frame_data_bvb, text='1) Выберите отчет по форме Минпросвещения', font=('Arial Bold', 14),
                           command=select_data_rmg)
    btn_bvb_rmg.pack(padx=10, pady=10)

    btn_bvb_students = Button(frame_data_bvb, text='2) Выберите сводный отчет по ученикам', font=('Arial Bold', 14),
                           command=select_data_bvb)
    btn_bvb_students.pack(padx=10, pady=10)


    btn_bvb_choose_end_folder = Button(frame_data_bvb, text='3) Выберите конечную папку',
                                       font=('Arial Bold', 14),
                                       command=select_end_folder_bvb
                                       )
    btn_bvb_choose_end_folder.pack(padx=10, pady=10)

    # Создаем кнопку слияния

    btn_bvb_process = Button(tab_bvb, text='4) Выполнить обработку',
                             font=('Arial Bold', 20),
                             command=processing_create_svod_bvb)
    btn_bvb_process.pack(padx=10, pady=10)


    """
    Разница двух таблиц
    """
    tab_diffrence = Frame(tab_control)
    tab_control.add(tab_diffrence, text='Разница 2 таблиц')

    diffrence_frame_description = LabelFrame(tab_diffrence)
    diffrence_frame_description.pack()

    lbl_hello_diffrence = Label(diffrence_frame_description,
                                text='Поиск отличий в двух таблицах\n'
                                     'ВАЖНО Количество строк и колонок в таблицах должно совпадать\n'
                                     'ВАЖНО Названия колонок в таблицах должны совпадать\n'
                                     'ПРИМЕЧАНИЯ\n'
                                     'Заголовок таблицы должен занимать только первую строку!\n'
                                     'Для корректной работы программы уберите из таблицы\n объединенные ячейки',
                                width=60)

    lbl_hello_diffrence.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_diffrence = resource_path('logo.png')
    img_diffrence = PhotoImage(file=path_to_img_diffrence)
    Label(diffrence_frame_description,
          image=img_diffrence, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_diffrence = LabelFrame(tab_diffrence, text='Подготовка')
    frame_data_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_diffrence = Button(frame_data_diffrence, text='1) Выберите файл с первой таблицей',
                                      font=('Arial Bold', 14),
                                      command=select_first_diffrence
                                      )
    btn_data_first_diffrence.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_first_sheet_name_diffrence = StringVar()
    # Описание поля
    label_first_sheet_name_diffrence = Label(frame_data_diffrence,
                                             text='2) Введите название листа, где находится первая таблица')
    label_first_sheet_name_diffrence.pack(padx=10, pady=10)
    # поле ввода имени листа
    first_sheet_name_entry_diffrence = Entry(frame_data_diffrence, textvariable=entry_first_sheet_name_diffrence,
                                             width=30)
    first_sheet_name_entry_diffrence.pack(ipady=5)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_diffrence = Button(frame_data_diffrence, text='3) Выберите файл со второй таблицей',
                                       font=('Arial Bold', 14),
                                       command=select_second_diffrence
                                       )
    btn_data_second_diffrence.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_second_sheet_name_diffrence = StringVar()
    # Описание поля
    label_second_sheet_name_diffrence = Label(frame_data_diffrence,
                                              text='4) Введите название листа, где находится вторая таблица')
    label_second_sheet_name_diffrence.pack(padx=10, pady=10)
    # поле ввода
    second__sheet_name_entry_diffrence = Entry(frame_data_diffrence, textvariable=entry_second_sheet_name_diffrence,
                                               width=30)
    second__sheet_name_entry_diffrence.pack(ipady=5)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_diffrence = Button(frame_data_diffrence, text='5) Выберите конечную папку',
                                      font=('Arial Bold', 14),
                                      command=select_end_folder_diffrence
                                      )
    btn_select_end_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_diffrence = Button(tab_diffrence, text='6) Обработать таблицы', font=('Arial Bold', 20),
                                   command=processing_diffrence
                                   )
    btn_data_do_diffrence.pack(padx=10, pady=10)






    # """
    # Создаем вкладку
    # """
    # tab_template = ttk.Frame(tab_control)
    # tab_control.add(tab_template, text='Подготовка данных\nПервая профессия')
    #
    # template_frame_description = LabelFrame(tab_template)
    # template_frame_description.pack()
    #
    # lbl_hello_template = Label(template_frame_description,
    #                                text='Центр опережающей профессиональной подготовки Республики Бурятия', width=60)
    # lbl_hello_template.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    #
    # # Картинка
    # path_to_img_template = resource_path('logo.png')
    # img_template = PhotoImage(file=path_to_img_template)
    # Label(template_frame_description,
    #       image=img_template, padx=10, pady=10
    #       ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)
    #
    # # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    # frame_data_template = LabelFrame(tab_template, text='Подготовка')
    # frame_data_template.pack(padx=10, pady=10)
    #
    # # Создаем кнопку Выбрать файл
    #
    # btn_template_first = Button(frame_data_template, text='1) Выберите файл Яндекс формы', font=('Arial Bold', 14),
    #                             command=select_data_yandex_first_prof)
    # btn_template_first.pack(padx=10, pady=10)
    #
    #
    #
    # btn_template_choose_end_folder = Button(frame_data_template, text='2) Выберите конечную папку',
    #                                         font=('Arial Bold', 14),
    #                                         command=select_result_yandex_first_prof_folder
    #                                         )
    # btn_template_choose_end_folder.pack(padx=10, pady=10)
    #
    # # Создаем кнопку слияния
    #
    # btn_template_process = Button(tab_template, text='3) Выполнить обработку',
    #                               font=('Arial Bold', 20),
    #                               command=processing_preparation_yandex_first_prof)
    # btn_template_process.pack(padx=10, pady=10)
    #
    #
    # """
    # Создаем вкладку для создания Сводов
    # """
    # tab_svod_first_prof = ttk.Frame(tab_control)
    # tab_control.add(tab_svod_first_prof, text='Сводка\nПервая профессия')
    #
    # svod_first_prof_frame_description = LabelFrame(tab_svod_first_prof)
    # svod_first_prof_frame_description.pack()
    #
    # lbl_hello_svod_first_prof = Label(svod_first_prof_frame_description,
    #                                   text='Центр опережающей профессиональной подготовки Республики Бурятия', width=60)
    # lbl_hello_svod_first_prof.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    #
    # # Картинка
    # path_to_img_svod_first_prof = resource_path('logo.png')
    # img_svod_first_prof = PhotoImage(file=path_to_img_svod_first_prof)
    # Label(svod_first_prof_frame_description,
    #       image=img_svod_first_prof, padx=10, pady=10
    #       ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)
    #
    # # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    # frame_data_svod_first_prof = LabelFrame(tab_svod_first_prof, text='Подготовка')
    # frame_data_svod_first_prof.pack(padx=10, pady=10)
    #
    # # Создаем кнопку Выбрать файл
    #
    # btn_svod_itog_list_first_prof = Button(frame_data_svod_first_prof, text='1) Выберите итоговый список',
    #                                    font=('Arial Bold', 14),
    #                                    command=select_data_itog_list_first_prof)
    # btn_svod_itog_list_first_prof.pack(padx=10, pady=10)
    #
    # btn_svod_est_first_prof = Button(frame_data_svod_first_prof, text='2) Выберите файл с оценками',
    #                                    font=('Arial Bold', 14),
    #                                    command=select_data_est_first_prof)
    # btn_svod_est_first_prof.pack(padx=10, pady=10)
    #
    # btn_svod_moodle_first_prof = Button(frame_data_svod_first_prof, text='3) Выберите файл с логинами Moodle',
    #                                    font=('Arial Bold', 14),
    #                                    command=select_data_moodle_first_prof)
    # btn_svod_moodle_first_prof.pack(padx=10, pady=10)
    #
    #
    #
    # btn_svod_first_prof_choose_end_folder = Button(frame_data_svod_first_prof, text='4) Выберите конечную папку',
    #                                                font=('Arial Bold', 14),
    #                                                command=select_result_svod_folder
    #                                                )
    # btn_svod_first_prof_choose_end_folder.pack(padx=10, pady=10)
    #
    # # Создаем кнопку слияния
    #
    # btn_svod_first_prof_process = Button(tab_svod_first_prof, text='5) Выполнить обработку',
    #                                      font=('Arial Bold', 20),
    #                                      command=processing_create_svod_first_prof)
    # btn_svod_first_prof_process.pack(padx=10, pady=10)
    #
















    # Создаем виджет для управления полосой прокрутки
    canvas.create_window((0, 0), window=tab_control, anchor="nw")

    # Конфигурируем холст для обработки скроллинга
    canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))
    scrollbar.pack(side="right", fill="y")

    # Вешаем событие скроллинга
    canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)
    window.mainloop()










