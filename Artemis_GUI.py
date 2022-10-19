"""
Скрипт для обработки создания отчетов по площадям леса
"""
import pandas as pd
import numpy as np
import os
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
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
    return ';'.join(x)


def check_unique(x):
    # Функция для нахождения разночтений в площади выделенного гектара
    temp_lst = x.split(';')
    temp_set = set(temp_lst)
    if 'nan' in temp_set:
        return 'Не заполнены значения площади лесотаксационного выдела!!!'
    else:
        return 'Все в порядке' if len(temp_set) == 1 else 'Ошибка!!!'


def main_check_unique(x):
    # Функция для проверки корректности заполнения площади выдела
    temp_str = ';'.join(x)  # Склеиваем все значения
    temp_lst = temp_str.split(';')  # Создаем список разбивая по ;
    temp_set = set(temp_lst)  # Превращаем в множество
    if len(temp_set) > 1:  # Если длина множества больше 1 то есть погрешности
        return 0
    elif 'nan' in temp_set:  # если есть нан то не заполнены площади выдела
        return 0
    else:  # Если все в порядке то возвращаем единственный элемент списка
        try:
            return float(temp_lst[0])
        except ValueError:
            messagebox.showerror('Артемида 1.1',
                                 f'При обработке значения {x} в столбце с числовыми данными возникла ошибка\n'
                                 f'Исправьте это значение и попробуйте заново')


def convert_to_float(x):
    """
    Функция для конвертирования строки в флоат при ошибке возвращает 0
    :param x: строка
    :return:
    """
    temp_str = x.strip()
    try:
        return float(temp_str)
    except ValueError:
        return 0

def clean_purpose_column(x):
    """
    Функция для извлечения значений из столбца целевого назначения для того чтобы можно было
    найти все значения равные 1 и сопоставить со значением в категории
    """
    temp_lst = x.split(';')  # Создаем список разделя строку по ;
    temp_set = set(temp_lst)  # Превращаем во множество

    if len(temp_set) == 1:
        temp_value = list(temp_set)[0]  # получаем единственное значение
        if temp_value == 'nan':
            return 0
        try:
            value_purpose = float(temp_value)  # конвертируем в число

            return value_purpose
        except ValueError:
            return 0
    else:
        return 0

def prepare_column_purpose_category(df,name_columns):
    """
    Функция для предобработки колонок с целевым назначением и категорией лесов
    Нужно очистить от пробелов, nan, сконвертировать во флоат,инт и снова в строку
    :param df: датафрейм содержащий в себе реестр
    :param name_columns: название обрабаытваемой колонки
    """
    try:
        # Приводим колонку к типу str чтобы очистить от лишних символов и заменить пустые вещи на нули
        df[name_columns] = df[name_columns].astype(str)
        df[name_columns] = df[name_columns].apply(lambda x: x.replace('nan', '0'))
        df[name_columns] = df[name_columns].apply(lambda x: x.replace(' ', '0'))
        df[name_columns] = df[name_columns].apply(lambda x: x.strip())
        # конвертируем во флоат, затем в инт чтобы потом в строке не было значений с дробной частью
        df[name_columns] = df[name_columns].astype(float)
        df[name_columns] = df[name_columns].astype(int)
        df[name_columns] = df[name_columns].astype(str)
    except KeyError as e:
        messagebox.showerror('Артемида 1.1',f'Не найдена колонка {e.args} Проверьте файл на наличие этой колонки')
    except ValueError as e:
        messagebox.showerror('Артемида 1.1', f'Возникла ошибка при обработке значения {e.args}\n'
                                             f'в колонках целевого назначения и категории должны быть только цифры!')


def processing_report_square_wood():
    """
    Фугкция для обработки данных
    :return:
    """
    try:
        df = pd.read_excel(file_data_xlsx, sheet_name='Реестр УПП', skiprows=6)

        # Удаляем лишние строки
        df = df.drop([0, 1], axis=0)

        # Заполняем незаполненные поля в стобце урочище
        df['Урочище '] = df['Урочище '].fillna('Название урочища не заполнено')

        # СОздаем проверочный файл для проверки правильности ввода плошади выдела
        check_df = df.copy(deep=True)

        # Меняем тип столбца на строку чтобы создать строку включающую в себя все значения разделенные ;
        check_df['Площадь лесотаксационного выдела, га'] = check_df['Площадь лесотаксационного выдела, га'].astype(str)
        #

        checked_pl = check_df.groupby(['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                                       'Номер лесотаксационного выдела']).agg(
            {'Площадь лесотаксационного выдела, га': combine})

        # Применяем функцию првоеряющую количество уникальных значений в столбце, если больше одного то значит есть ошибка в данных
        checked_pl['Контроль совпадения площади выдела'] = checked_pl['Площадь лесотаксационного выдела, га'].apply(
            check_unique)

        # переименовывам колонку
        checked_pl.rename(
            columns={'Площадь лесотаксационного выдела, га': 'Все значения площади для указанного выдела'},
            inplace=True)
        # Извлекаем индексы в колонки
        checked_pl = checked_pl.reset_index()

        # Получаем текущую дату
        current_time = time.strftime('%H_%M_%S %d.%m.%Y')
        # Сохраняем отчет
        # Для того чтобы увеличить ширину колонок для удобства чтения используем openpyxl
        wb = openpyxl.Workbook()  # Создаем объект
        # Записываем результаты
        for row in dataframe_to_rows(checked_pl, index=False, header=True):
            wb['Sheet'].append(row)

        # Форматирование итоговой таблицы
        # Ширина колонок
        wb['Sheet'].column_dimensions['A'].width = 15
        wb['Sheet'].column_dimensions['B'].width = 20
        wb['Sheet'].column_dimensions['C'].width = 36
        wb['Sheet'].column_dimensions['F'].width = 15
        wb['Sheet'].column_dimensions['G'].width = 15
        # Перенос строк для заголовков
        wb['Sheet']['D1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['E1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['F1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['G1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['H1'].alignment = Alignment(wrap_text=True)

        wb.save(
            f'{path_to_end_folder}/Проверка правильности ввода площадей лесотаксационного выдела {current_time}.xlsx')

        # Основной отчет
        # Готовим колонки к группировке
        df['Площадь лесотаксационного выдела, га'] = df['Площадь лесотаксационного выдела, га'].astype(str)

        df['Площадь лесотаксационного выдела, га'] = df['Площадь лесотаксационного выдела, га'].apply(
            lambda x: x.replace(',', '.'))

        df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
            'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].astype(str)

        df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
            'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].apply(
            lambda x: x.replace(',', '.'))

        df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
            'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].apply(convert_to_float)

        # Группируем
        group_df = df.groupby(['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                               'Номер лесотаксационного выдела']).agg(
            {'Площадь лесотаксационного выдела, га': main_check_unique,
             'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га': 'sum'})

        # переименовываем колонку
        group_df.rename(columns={
            'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га': 'Используемая площадь лесотаксационного выдела, га'},
            inplace=True)

        # Извлекаем индексы в колонки
        group_df = group_df.reset_index()

        group_df['Площадь лесотаксационного выдела, га'] = group_df['Площадь лесотаксационного выдела, га'].astype(
            float)

        # Округляем до 3 знаков для корректного сравнения
        group_df['Площадь лесотаксационного выдела, га'] = np.round(group_df['Площадь лесотаксационного выдела, га'],
                                                                    decimals=3)
        group_df['Используемая площадь лесотаксационного выдела, га'] = np.round(
            group_df['Используемая площадь лесотаксационного выдела, га'], decimals=3)

        # Создаем колонку для контроля
        group_df['Контроль площади используемого надела'] = group_df['Площадь лесотаксационного выдела, га'] < group_df[
            'Используемая площадь лесотаксационного выдела, га']

        group_df['Контроль площади используемого надела'] = group_df['Контроль площади используемого надела'].apply(
            lambda x: 'Превышение используемой площади выдела!!!' if x is True else 'Все в порядке')

        # Изменяем состояние колонки если площадь всего выдела равна 0
        group_df['Контроль правильности ввода площади лесотаксационного выдела'] = group_df[
            'Площадь лесотаксационного выдела, га'].apply(
            lambda
                x: 'Площадь лесотаксационного выдела равна нулю или  обнаружены разные значения площади выдела !!!' if x == 0 else 'Значения лесотаксационного выдела не отличаются друг от друга')

        # Сохраняем отчет
        # Для того чтобы увеличить ширину колонок для удобства чтения используем openpyxl
        wb = openpyxl.Workbook()  # Создаем объект
        # Записываем результаты
        for row in dataframe_to_rows(group_df, index=False, header=True):
            wb['Sheet'].append(row)

        # Форматирование итоговой таблицы
        # Ширина колонок
        wb['Sheet'].column_dimensions['A'].width = 15
        wb['Sheet'].column_dimensions['B'].width = 20
        wb['Sheet'].column_dimensions['C'].width = 36
        wb['Sheet'].column_dimensions['F'].width = 15
        wb['Sheet'].column_dimensions['G'].width = 15
        wb['Sheet'].column_dimensions['H'].width = 20
        # Перенос строк для заголовков
        wb['Sheet']['D1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['E1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['F1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['G1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['H1'].alignment = Alignment(wrap_text=True)

        wb.save(f'{path_to_end_folder}/Контроль используемых площадей лесотаксационных выделов от {current_time}.xlsx')



    except KeyError as e:
        messagebox.showerror('Артемида 1.1',
                             f'Не найдена колонка или лист {e.args}\nДанные в файле должны находиться на листе с названием Реестр УПП'
                             f'\nКолонки 1-8 должны иметь названия: Лесничество,Участковое лесничество,Урочище ,Номер лесного квартала,\n'
                             f'Номер лесотаксационного выдела,Площадь лесотаксационного выдела, га,Обозначение части лесотаксационного выдела (лесопатологического выдела), га ,'
                             f'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га')
    except NameError:
        messagebox.showerror('Артемида 1.1', f'Выберите файл с данными и конечную папку')
    except PermissionError:
        messagebox.showerror('Артемида 1.1', f'Закройте файлы с созданными раньше отчетами!!!')
    else:
        messagebox.showinfo('Артемида 1.1', 'Работа программы успешно завершена!!!')

def proccessing_report_purpose_category():
    """
    Функция для обработки реестра УПП, находит некорректно заполненные графы целевого назначения
    лесов и категории(графы 12 и графы 13)
    """
    try:
        df = pd.read_excel(file_data_xlsx, skiprows=6,
                           usecols=['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                                    'Номер лесотаксационного выдела'
                               , 'Целевое назначение лесов ', 'Категория защитных лесов (код) '])

        # Удаляем лишние строки
        df = df.drop([0, 1], axis=0)
        # заполняем пропущенные места
        df['Урочище '] = df['Урочище '].fillna('Название урочища не заполнено')

        # Меняем тип столбца на строку чтобы создать строку включающую в себя все значения разделенные ;заменяем нан на нули и очищаем от пробельных символов
        prepare_column_purpose_category(df, 'Целевое назначение лесов ')
        prepare_column_purpose_category(df, 'Категория защитных лесов (код) ')

        checked_pl = df.groupby(['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                                 'Номер лесотаксационного выдела']).agg(
            {'Целевое назначение лесов ': combine, 'Категория защитных лесов (код) ': combine})

        # Извлекаем индекс
        out_df = checked_pl.reset_index()

        # Применяем функцию првоеряющую количество уникальных значений в столбце, если больше одного то значит есть ошибка в данных
        out_df['Контроль правильности заполнения целевого назначения лесов'] = out_df['Целевое назначение лесов '].apply(
            check_unique)
        out_df['Контроль одинаковости заполнения категории защитных лесов'] = out_df[
            'Категория защитных лесов (код) '].apply(
            check_unique)

        out_df['Контроль назначения лесов'] = out_df['Целевое назначение лесов '].apply(clean_purpose_column)

        out_df['Контроль назначения лесов'] = out_df['Контроль назначения лесов'].astype(
            int)  # Приводим на всякий случай к инту

        out_df['Контроль категории защитных лесов'] = out_df['Категория защитных лесов (код) '].apply(clean_purpose_column)
        out_df['Контроль категории защитных лесов'] = out_df['Контроль категории защитных лесов'].astype(
            int)  # Приводим на всякий случай к инту

        out_df.rename(columns={'Целевое назначение лесов ': 'Показатели целевого назначения для данного выдела',
                               'Категория защитных лесов (код) ': 'Показатели категории защитных лесов для данного выдела'},
                      inplace=True)

        out_df['Итоговый контроль защитных лесов'] = (out_df['Контроль назначения лесов'] == 1) & (
                out_df['Контроль категории защитных лесов'] == 0)

        out_df['Итоговый контроль защитных лесов'] = out_df['Итоговый контроль защитных лесов'].apply(
            lambda x: 'Ошибка, проверьте целевое назначение или категорию защитных лесов' if x == True else 'Все в порядке')

        # Получаем текущую дату
        current_time = time.strftime('%H_%M_%S %d.%m.%Y')
        # Сохраняем отчет
        # Для того чтобы увеличить ширину колонок для удобства чтения используем openpyxl
        wb = openpyxl.Workbook()  # Создаем объект
        # Записываем результаты
        for row in dataframe_to_rows(out_df, index=False, header=True):
            wb['Sheet'].append(row)

        # Форматирование итоговой таблицы
        # Ширина колонок
        wb['Sheet'].column_dimensions['A'].width = 15
        wb['Sheet'].column_dimensions['B'].width = 20
        wb['Sheet'].column_dimensions['C'].width = 36
        wb['Sheet'].column_dimensions['F'].width = 15
        wb['Sheet'].column_dimensions['G'].width = 15
        wb['Sheet'].column_dimensions['H'].width = 15
        wb['Sheet'].column_dimensions['I'].width = 15
        wb['Sheet'].column_dimensions['J'].width = 15
        wb['Sheet'].column_dimensions['K'].width = 15
        wb['Sheet'].column_dimensions['L'].width = 15
        # Перенос строк для заголовков
        wb['Sheet']['F1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['G1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['H1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['I1'].alignment = Alignment(wrap_text=True)
        wb['Sheet']['L1'].alignment = Alignment(wrap_text=True)

        wb.save(
            f'{path_to_end_folder}/Проверка правильности ввода целевого назначения лесов и категории защитных лесов {current_time}.xlsx')
    except ValueError as e:
        messagebox.showerror('Артемида 1.1',
                             f'Не найдена колонка или лист {e.args}\nДанные в файле должны находиться на листе с названием Реестр УПП'
                             f'\nВ файле должны быть колонки: Лесничество,Участковое лесничество,Урочище ,Номер лесного квартала,\n'
                             f'Номер лесотаксационного выдела,Целевое назначение лесов ,Категория защитных лесов (код) ')
    except NameError:
        messagebox.showerror('Артемида 1.1', f'Выберите файл с данными и конечную папку')
    except PermissionError:
        messagebox.showerror('Артемида 1.1', f'Закройте файлы с созданными раньше отчетами!!!')
    else:
        messagebox.showinfo('Артемида 1.1', 'Работа программы успешно завершена!!!')


if __name__ == '__main__':
    window = Tk()
    window.title('Артемида 1.1')
    window.geometry('760x560+600+200')
    window.resizable(False, False)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)


    # Создаем вкладку обработки данных по площадям выделов
    tab_report_square = ttk.Frame(tab_control)
    tab_control.add(tab_report_square, text='Контроль площадей\nвыделов')
    tab_control.pack(expand=1, fill='both')

    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_report_square,
                      text='Центр защиты леса Республики Бурятия\n'
                           'Контроль соответствия площадей выделов,\nправильности заполнения площадей выделов ')
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

    # Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_report_square, text='3) Обработать данные', font=('Arial Bold', 20),
                                  command=processing_report_square_wood
                                  )
    btn_proccessing_data.grid(column=0, row=4, padx=10, pady=10)

    # Создаем вкладку обработки данных по площадям выделов
    tab_report_purpose_category = ttk.Frame(tab_control)
    tab_control.add(tab_report_purpose_category, text='Контроль назначения и\n категории защитности')

    # Создаем метку для описания назначения программы
    lbl_hello_purpose = Label(tab_report_purpose_category,
                      text='Центр защиты леса Республики Бурятия\n'
                           'Контроль соответствия целевого назначения лесов\nи категории защитности')
    lbl_hello_purpose.grid(column=0, row=0, padx=10, pady=25)


    # Картинка
    path_to_img = resource_path('logo.png')

    img_purpose = PhotoImage(file=path_to_img)
    Label(tab_report_purpose_category,
          image=img_purpose
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data_purpose = Button(tab_report_purpose_category, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                             command=select_file_data_xlsx
                             )
    btn_choose_data_purpose.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_purpose = Button(tab_report_purpose_category, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder_purpose.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_data_purpose = Button(tab_report_purpose_category, text='3) Обработать данные', font=('Arial Bold', 20),
                                  command=proccessing_report_purpose_category
                                  )
    btn_proccessing_data_purpose.grid(column=0, row=4, padx=10, pady=10)


    window.mainloop()
