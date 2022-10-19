"""
Проверка Соответствие по целевому назначению и категории защитности
"""
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import time

pd.options.mode.chained_assignment = None  # default='warn'


def combine(x):
    # Функция для группировки всех значений в строку разделенную ;
    return ';'.join(x)


def check_unique(x):
    # Функция для нахождения разночтений в площади выделенного гектара
    # создаем список разделяя по точке с запятой
    temp_lst = x.split(';')
    # Создаем множество оставляя только уникальные значения
    temp_set = set(temp_lst)
    return 'Значения совпадают' if len(temp_set) == 1 else 'Ошибка!!! Значения не совпадают'


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


path_reest_upp = 'data/Соответствие по целевому назначению и категории защитности/2022-10-27_64_Реестр УПП с дополнительными колонками..xlsx'
path_to_end_folder = 'data/'

df = pd.read_excel(path_reest_upp, skiprows=6,usecols=['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                            'Номер лесотаксационного выдела'
                       , 'Целевое назначение лесов ', 'Категория защитных лесов (код) '])

# Удаляем лишние строки
df = df.drop([0, 1], axis=0)
# заполняем пропущенные места
df['Урочище '] = df['Урочище '].fillna('Название урочища не заполнено')


# Меняем тип столбца на строку чтобы создать строку включающую в себя все значения разделенные ;заменяем нан на нули и очищаем от пробельных символов
df['Целевое назначение лесов '] = df['Целевое назначение лесов '].astype(str)
df['Целевое назначение лесов '] = df['Целевое назначение лесов '].apply(lambda x: x.replace('nan', '0'))
df['Целевое назначение лесов '] = df['Целевое назначение лесов '].apply(lambda x: x.replace(' ', '0'))
df['Целевое назначение лесов '] = df['Целевое назначение лесов '].apply(lambda x: x.strip())

df['Категория защитных лесов (код) '] = df['Категория защитных лесов (код) '].astype(str)
df['Категория защитных лесов (код) '] = df['Категория защитных лесов (код) '].apply(lambda x: x.replace('nan', '0'))
df['Категория защитных лесов (код) '] = df['Категория защитных лесов (код) '].apply(lambda x: x.replace(' ', '0'))
df['Категория защитных лесов (код) '] = df['Категория защитных лесов (код) '].apply(lambda x: x.strip())

checked_pl = df.groupby(['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                         'Номер лесотаксационного выдела']).agg(
    {'Целевое назначение лесов ': combine, 'Категория защитных лесов (код) ': combine})

# Извлекаем индекс
out_df = checked_pl.reset_index()

# Применяем функцию првоеряющую количество уникальных значений в столбце, если больше одного то значит есть ошибка в данных
out_df['Контроль правильности заполнения целевого назначения лесов'] = out_df['Целевое назначение лесов '].apply(
    check_unique)
out_df['Контроль одинаковости заполнения категории защитных лесов'] = out_df['Категория защитных лесов (код) '].apply(
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

out_df.head()

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

# out_df.to_excel('Тест.xlsx',index=False)









