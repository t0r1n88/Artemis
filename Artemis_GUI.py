"""
Скрипт для обработки создания отчетов по площадям леса
"""
import pandas as pd
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time
# pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def select_end_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder
    path_to_end_folder = filedialog.askdirectory()

def select_file_data_xlsx():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_xlsx
    # Получаем путь к файлу
    file_data_xlsx = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def combine(x):
    # Функция для группировки всех значений в строку разделенную ;
    return  ';'.join(x)

def check_unique(x):
    # Функция для нахождения разночтений в площади выделенного гектара
    temp_lst = x.split(';')
    temp_set = set(temp_lst)
    if'nan' in temp_set:
        return 'Не заполнены значения площади лесотаксационного выдела!!!'
    else:
        return 'Площади совпадают' if len(temp_set) == 1 else 'Ошибка!!! Площади не совпадают'



def processing_report_square_wood():
    """
    Фугкция для обработки данных
    :return:
    """
    try:
        df = pd.read_excel(file_data_xlsx, sheet_name='Реестр УПП', skiprows=6)
        # Удаляем лишние строки
        df = df.drop([0, 1], axis=0)

        # СОздаем проверочный файл для проверки правильности ввода плошади выдела
        check_df = df.copy()

        # Меняем тип столбца на строку чтобы создать строку включающую в себя все значения разделенные ;
        check_df['Площадь лесотаксационного выдела, га'] = check_df['Площадь лесотаксационного выдела, га'].astype(str)

        checked_pl = check_df.groupby(['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                                       'Номер лесотаксационного выдела']).agg(
            {'Площадь лесотаксационного выдела, га': combine})

        # Применяем функцию првоеряющую количество уникальных значений в столбце, если больше одного то значит есть ошибка в данных
        checked_pl['Контроль совпадения площади выдела'] = checked_pl['Площадь лесотаксационного выдела, га'].apply(
            check_unique)

        # переименовывам колонку
        checked_pl.rename(columns={'Площадь лесотаксационного выдела, га': 'Все значения площади для указанного выдела'},
                          inplace=True)
        # Извлекаем индексы в колонки
        checked_pl = checked_pl.reset_index()
        # Заполняем nan в колонке со значениями площади
        # checked_pl['Контроль совпадения площади выдела'] = checked_pl['Все значения площади для указанного выдела'].apply(lambda x:'Не заполнены значения площади!!!' if x == 'nan' else x)

        # Получаем текущую дату
        current_time = time.strftime('%d.%m.%Y')
        # Сохраняем отчет

        checked_pl.to_excel(f'{path_to_end_folder}/Отчет Контроль совпадения площади выдела от {current_time}.xlsx', index=False)

        # Основной отчет
        # Готовим колонки к группировке
        df['Площадь лесотаксационного выдела, га'] = df['Площадь лесотаксационного выдела, га'].astype(str)

        df['Площадь лесотаксационного выдела, га'] = df['Площадь лесотаксационного выдела, га'].apply(
            lambda x: x.replace(',', '.'))

        df['Площадь лесотаксационного выдела, га'] = df['Площадь лесотаксационного выдела, га'].astype(float)

        df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
            'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].astype(str)

        df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
            'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].apply(
            lambda x: x.replace(',', '.'))

        df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
            'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].astype(float)

        # Группируем
        group_df = df.groupby(['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                               'Номер лесотаксационного выдела']).agg({'Площадь лесотаксационного выдела, га': 'sum',
                                                                       'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га': 'sum'})

        group_df.rename(columns={
            'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га': 'Используемая площадь лесотаксационного выдела, га'},
                        inplace=True)

        # Извлекаем индексы в колонки
        group_df = group_df.reset_index()

        group_df['Площадь лесотаксационного выдела, га'] = group_df['Площадь лесотаксационного выдела, га'].astype(float)
        group_df['Используемая площадь лесотаксационного выдела, га'] = group_df[
            'Используемая площадь лесотаксационного выдела, га'].astype(float)

        group_df['Контроль площади используемого надела'] = group_df['Площадь лесотаксационного выдела, га'] < group_df[
            'Используемая площадь лесотаксационного выдела, га']

        group_df['Контроль площади используемого надела'] = group_df['Контроль площади используемого надела'].apply(
            lambda x: 'Превышение используемой площади выдела!!!' if x is True else 'Все в порядке')

        group_df.to_excel(f'{path_to_end_folder}/Отчет о площадях выделов от {current_time}.xlsx', index=False)
    except KeyError as e:
        messagebox.showerror('Артемида 1.0',f'Не найдена колонка или лист {e.args}\nДанные в файле должны находиться на листе с названием Реестр УПП'
                                            f'\nКолонки 1-8 должны иметь названия: Лесничество,Участковое лесничество,Урочище ,Номер лесного квартала,\n'
                                            f'Номер лесотаксационного выдела,Площадь лесотаксационного выдела, га,Обозначение части лесотаксационного выдела (лесопатологического выдела), га ,'
                                            f'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га')
    except NameError:
        messagebox.showerror('Артемида 1.0', f'Выберите файл с данными и конечную папку')
    except PermissionError:
        messagebox.showerror('Артемида 1.0', f'Закройте файлы с созданными раньше отчетами!!!')


if __name__ == '__main__':
    window = Tk()
    window.title('Артемида 1.0')
    window.geometry('700x560+600+200')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных
    tab_report_square = ttk.Frame(tab_control)
    tab_control.add(tab_report_square, text='Отчеты по площадям выделов')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_report_square,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nЦентр защиты леса Республики Бурятия')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')

    img = PhotoImage(file=path_to_img)
    Label(tab_report_square,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data = Button(tab_report_square, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                             command=select_file_data_xlsx
                             )
    btn_choose_data.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_report_square, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)

    #Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_report_square, text='3) Обработать данные', font=('Arial Bold', 20),
                                  command=processing_report_square_wood
                                  )
    btn_proccessing_data.grid(column=0, row=4, padx=10, pady=10)

    window.mainloop()