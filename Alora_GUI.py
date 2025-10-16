from  first_prof import processing_data_first_prof # для обработки файла первой профессии
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


def processing_preparation_yandex_first_prof():
    """
    Функция для генерации документов
    """
    try:
        processing_data_first_prof(data_yandex_first_prof,path_to_first_prof_end_folder)

    except NameError:
        messagebox.showerror('',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')



if __name__ == '__main__':
    window = Tk()
    window.title('Алора ver 1.2')
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
    Создаем вкладку 
    """
    tab_template = ttk.Frame(tab_control)
    tab_control.add(tab_template, text='Подготовка данных\nПервая профессия')

    template_frame_description = LabelFrame(tab_template)
    template_frame_description.pack()

    lbl_hello_template = Label(template_frame_description,
                                   text='Центр опережающей профессиональной подготовки Республики Бурятия', width=60)
    lbl_hello_template.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_template = resource_path('logo.png')
    img_template = PhotoImage(file=path_to_img_template)
    Label(template_frame_description,
          image=img_template, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_template = LabelFrame(tab_template, text='Подготовка')
    frame_data_template.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать файл

    btn_template_first = Button(frame_data_template, text='1) Выберите файл Яндекс формы', font=('Arial Bold', 14),
                                command=select_data_yandex_first_prof)
    btn_template_first.pack(padx=10, pady=10)



    btn_template_choose_end_folder = Button(frame_data_template, text='2) Выберите конечную папку',
                                            font=('Arial Bold', 14),
                                            command=select_result_yandex_first_prof_folder
                                            )
    btn_template_choose_end_folder.pack(padx=10, pady=10)

    # Создаем кнопку слияния

    btn_template_process = Button(tab_template, text='3) Выполнить обработку',
                                  font=('Arial Bold', 20),
                                  command=processing_preparation_yandex_first_prof)
    btn_template_process.pack(padx=10, pady=10)










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










