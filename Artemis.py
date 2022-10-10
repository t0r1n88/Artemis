"""
Скрипт для обработки создания отчетов по площадям леса
"""
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import time
import numpy as np

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
        return 'Площади совпадают' if len(temp_set) == 1 else 'Ошибка!!! Площади лесотаксационного выдела не совпадают'

def main_check_unique(x):
    # Функция для проверки корректности заполнения площади выдела
    temp_str = ';'.join(x) # Склеиваем все значения
    temp_lst = temp_str.split(';') # Создаем список разбивая по ;
    temp_set = set(temp_lst) # Превращаем в множество
    if len(temp_set) > 1: # Если длина множества больше 1 то есть погрешности
        return 0
    elif'nan' in temp_set:# если есть нан то не заполнены площади выдела
        return 0
    else:# Если все в порядке то возвращаем единственный элемент списка
        return float(temp_lst[0])


file_data_xlsx = 'data/BLPK.xlsx'
path_to_end_folder = 'data/'

df = pd.read_excel(file_data_xlsx,sheet_name='Реестр УПП',skiprows=6)

# Удаляем лишние строки
df = df.drop([0,1],axis=0)

# Заполняем незаполненные поля в стобце урочище
df['Урочище '] = df['Урочище '].fillna('Название урочища не заполнено')


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
current_time = time.strftime('%H_%M_%S %d.%m.%Y')
# Сохраняем отчет
# Для того чтобы увеличить ширину колонок для удобства чтения используем openpyxl
wb = openpyxl.Workbook() # Создаем объект
# Записываем результаты
for row in dataframe_to_rows(checked_pl,index=False,header=True):
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


wb.save(f'{path_to_end_folder}/Проверка правильности ввода площадей лесотаксационного выдела {current_time}.xlsx')

# Основной отчет
# Готовим колонки к группировке
df['Площадь лесотаксационного выдела, га'] = df['Площадь лесотаксационного выдела, га'].astype(str)

df['Площадь лесотаксационного выдела, га'] = df['Площадь лесотаксационного выдела, га'].apply(
    lambda x: x.replace(',', '.'))

# df['Площадь лесотаксационного выдела, га'] = df['Площадь лесотаксационного выдела, га'].astype(float)

df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
    'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].astype(str)

df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
    'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].apply(
    lambda x: x.replace(',', '.'))

df['Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'] = df[
    'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га'].astype(float)



# Группируем
group_df = df.groupby(['Лесничество', 'Участковое лесничество', 'Урочище ', 'Номер лесного квартала',
                       'Номер лесотаксационного выдела']).agg({'Площадь лесотаксационного выдела, га': main_check_unique,
                                                               'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га': 'sum'})

# переименовываем колонку
group_df.rename(columns={
    'Площадь лесотаксационного выдела или его части (лесопатологического выдела), га': 'Используемая площадь лесотаксационного выдела, га'},
                inplace=True)

# Извлекаем индексы в колонки
group_df = group_df.reset_index()

group_df['Площадь лесотаксационного выдела, га'] = group_df['Площадь лесотаксационного выдела, га'].astype(float)

# Округляем до 3 знаков для корректного сравнения
group_df['Площадь лесотаксационного выдела, га'] = np.round(group_df['Площадь лесотаксационного выдела, га'],decimals=3)
group_df['Используемая площадь лесотаксационного выдела, га'] = np.round(group_df['Используемая площадь лесотаксационного выдела, га'],decimals=3)




# Создаем колонку для контроля
group_df['Контроль площади используемого надела'] = group_df['Площадь лесотаксационного выдела, га'] < group_df[
    'Используемая площадь лесотаксационного выдела, га']

group_df['Контроль площади используемого надела'] = group_df['Контроль площади используемого надела'].apply(
    lambda x: 'Превышение используемой площади выдела!!!' if x is True else 'Все в порядке')

# Изменяем состояние колонки если площадь всего выдела равна 0
group_df['Контроль правильности ввода площади лесотаксационного выдела'] = group_df['Площадь лесотаксационного выдела, га'].apply(
    lambda x: 'Площадь лесотаксационного выдела равна нулю или  обнаружены разные значения площади выдела !!!' if x == 0 else 'Значения лесотаксационного выдела не отличаются друг от друга')



# Сохраняем отчет
# Для того чтобы увеличить ширину колонок для удобства чтения используем openpyxl
wb = openpyxl.Workbook() # Создаем объект
# Записываем результаты
for row in dataframe_to_rows(group_df,index=False,header=True):
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




