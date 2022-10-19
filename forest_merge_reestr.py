"""
Проверка наличия данных из ведомости лесов в реестре УПП с отображением
"""
import pandas as pd
import openpyxl
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import time

pd.options.mode.chained_assignment = None  # default='warn'


def convert_columns_to_str(df, number_columns):
    """
    Функция для конвертации указанных столбцов в строковый тип и очистки от пробельных символов в начале и конце
    """

    for column in number_columns:  # Перебираем список нужных колонок
        df.iloc[:, column] = df.iloc[:, column].astype(str)
        # Очищаем колонку от пробельных символов с начала и конца
        df.iloc[:, column] = df.iloc[:, column].apply(lambda x: x.strip())
        df.iloc[:, column] = df.iloc[:, column].apply(lambda x: x.replace(' ', ''))


def convert_params_columns_to_int(lst):
    """
    Функция для конвератации значений колонок которые нужно обработать.
    Очищает от пустых строк, чтобы в итоге остался список из чисел в формате int
    """
    out_lst = []  # Создаем список в который будем добавлять только числа
    for value in lst:  # Перебираем список
        try:
            # Обрабатываем случай с нулем, для того чтобы после приведения к питоновскому отсчету от нуля не получилась колонка с номером -1
            number = int(value)
            if number != 0:
                out_lst.append(value)  # Если конвертирования прошло без ошибок то добавляем
            else:
                continue
        except:  # Иначе пропускаем
            continue
    return out_lst


file_params = 'data/params.xlsx'
path_to_end_folder = 'data/'
params = pd.read_excel(file_params, header=None,
                       keep_default_na=False)  # получаем файл с порядковыми номерами колонок которые нужно сравнивать
file_reestr_upp = 'data/Сравнение УПП с другими ведомостями на наличие участков или их отсутствие/2022-10-27_64_Реестр УПП с дополнительными колонками..xlsx'
file_statement_on_reestr = 'data/Сравнение УПП с другими ведомостями на наличие участков или их отсутствие/Ведомость.xlsx'

# Преврашаем каждую колонку в список
params_first_columns = params[0].tolist()
params_second_columns = params[1].tolist()

# Конвертируем в инт заодно проверяя корректность введенных данных
int_params_first_columns = convert_params_columns_to_int(params_first_columns)
int_params_second_columns = convert_params_columns_to_int(params_second_columns)

# Отнимаем 1 от каждого значения чтобы привести к питоновским индексам
int_params_first_columns = list(map(lambda x: x - 1, int_params_first_columns))
int_params_second_columns = list(map(lambda x: x - 1, int_params_second_columns))

# Считываем из файлов только те колонки по которым будет вестись сравнение
first_df = pd.read_excel(file_reestr_upp,
                         skiprows=6, usecols=int_params_first_columns, keep_default_na=False)
second_df = pd.read_excel(file_statement_on_reestr, usecols=int_params_second_columns, keep_default_na=False)
# Конвертируем нужные нам колонки в str
convert_columns_to_str(first_df, int_params_first_columns)
convert_columns_to_str(second_df, int_params_second_columns)

# Создаем в каждом датафрейме колонку с айди путем склеивания всех нужных колонок в одну строку
first_df['ID'] = first_df.iloc[:, int_params_first_columns].sum(axis=1)
second_df['ID'] = second_df.iloc[:, int_params_second_columns].sum(axis=1)

# Обрабатываем дубликаты


first_df.drop_duplicates(subset=['ID'], keep='last', inplace=True)  # Удаляем дубликаты из датафрейма

second_df.drop_duplicates(subset=['ID'], keep='last', inplace=True)  # Удаляем дубликаты из датафрейма

# Создаем документ
wb = openpyxl.Workbook()
# создаем листы
ren_sheet = wb['Sheet']
ren_sheet.title = 'Итог'

# Создаем датафрейм
itog_df = pd.merge(first_df, second_df, how='outer', left_on=['ID'], right_on=['ID'],
                   indicator=True)

# Отфильтровываем значения both,right
out_df = itog_df[(itog_df['_merge'] == 'both') | (itog_df['_merge'] == 'right_only')]

out_df.rename(columns={'_merge': 'Присутствие в реестре УПП'}, inplace=True)

out_df['Присутствие в реестре УПП'] = out_df['Присутствие в реестре УПП'].apply(
    lambda x: 'Имеется в реестре' if x == 'both' else 'Отсутствует в реестре')

# Получаем текущую дату
current_time = time.strftime('%H_%M_%S %d.%m.%Y')
# Сохраняем отчет
# Для того чтобы увеличить ширину колонок для удобства чтения используем openpyxl
wb = openpyxl.Workbook() # Создаем объект
# Записываем результаты
for row in dataframe_to_rows(out_df,index=False,header=True):
    wb['Sheet'].append(row)

# Форматирование итоговой таблицы
# Ширина колонок
wb['Sheet'].column_dimensions['A'].width = 15
wb['Sheet'].column_dimensions['B'].width = 20
wb['Sheet'].column_dimensions['C'].width = 10
wb['Sheet'].column_dimensions['F'].width = 36
wb['Sheet'].column_dimensions['L'].width = 20
# Перенос строк для заголовков
wb['Sheet']['D1'].alignment = Alignment(wrap_text=True)
wb['Sheet']['E1'].alignment = Alignment(wrap_text=True)
wb['Sheet']['F1'].alignment = Alignment(wrap_text=True)
wb['Sheet']['G1'].alignment = Alignment(wrap_text=True)
wb['Sheet']['H1'].alignment = Alignment(wrap_text=True)

wb.save(f'{path_to_end_folder}/Сравнение УПП с другими ведомостями на наличие участков или их отсутствие от  {current_time}.xlsx')
