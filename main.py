from datetime import timedelta
from datetime import datetime as dt
from typing import AnyStr

import xlrd

FILE = xlrd.open_workbook('task.xlsx')

SHEET = FILE.sheet_by_index(1)  # 1 - номер листа в документе

# Область захватываемых строк включительно. Строка 1 в Excel = Строка 0 в выдаче
RANGE = 2, 1002


"""
    Я ПРОЧИТАЛ ЭТО :)
    В решении я получил исключение ValueError: unconverted data remains: 
    "отметьте в решении, если вы прочитали это"
    
    Я удалил этот текст из таблицы. Но я также сохранил (закомментировал) функцию, 
    которая работает по-другому и обрабатывает этот столбец без ошибки.
    
    Написание кода у меня заняло 2.5 часа, тесты и комментирование - ещё 30 минут
    
    Автор: Илья Липчанский, https://t.me/lchansky
"""


def count_num1():
    column = 1
    count = 0
    for cell in range(*RANGE):
        if SHEET.cell_value(cell, column) % 2 == 0:
            count += 1
    print(f'Задание 1. Количество чётных чисел: {count}')


def count_num2():
    column = 2
    values = (
        int(SHEET.cell_value(cell, column)) for cell in range(*RANGE)
    )
    #  Тут в генераторе присваивается 1 на каждое простое число, и считается сумма этих единичек
    count = sum(1 for _ in filter(prime_num_check, values))
    print(f'Задание 2. Количество простых чисел: {count}')


def count_num3():
    column = 3
    values = (
        str_to_float(SHEET.cell_value(cell, column))
        for cell in range(*RANGE)
    )
    count = sum(1 for _ in filter(lambda x: x < 0.5, values))
    print(f'Задание 3. Количество простых чисел: {count}')


def count_date1():
    column = 4
    date_format = '%a %b %d %H:%M:%S %Y'
    values = (
        dt.strptime(SHEET.cell_value(cell, column), date_format).weekday()
        for cell in range(*RANGE)
    )
    count = sum(1 for value in values if value == 1)  # 0 - Monday, 1 - Tuesday etc.
    print(f'Задание 4. Количество вторников: {count}')


# def count_date1():
#     column = 4
#     values = (
#         SHEET.cell_value(cell, column).split(' ')[0]
#         for cell in range(*RANGE)
#     )
#     count = sum(1 for _ in filter(lambda x: x == 'Tue', values))
#     print(f'Задание 4. Количество вторников: {count}')


def count_date2():
    column = 5
    date_format = '%Y-%m-%d %H:%M:%S.%f'
    values = (
        dt.strptime(SHEET.cell_value(cell, column), date_format).weekday()
        for cell in range(*RANGE)
    )
    count = sum(1 for value in values if value == 1)  # 0 - Monday, 1 - Tuesday etc.
    print(f'Задание 5. Количество вторников: {count}')


def count_date3():
    column = 6
    count = 0
    date_format = '%m-%d-%Y'
    dates = (
        dt.strptime(SHEET.cell_value(cell, column), date_format)
        for cell in range(*RANGE)
    )
    for d in dates:
        if d.weekday() == 1 and (d + timedelta(7)).month != d.month:
            count += 1
    print(f'Задание 6. Количество последних вторников месяца: {count}')


def prime_num_check(num):
    """Принимает число num, возвращает True если оно простое"""
    if not isinstance(num, int) or num <= 1:
        return False
    # Сразу сокращаем диапазон поиска делителей пополам, т.к. 2 это минимальный делитель.
    # Начинаем с двойки, т.к. 1 - делитель у всех.
    for i in range(2, num // 2 + 1):
        if num % i == 0:  # Если найдётся хоть один делитель - сразу возвращаем False
            return False
    return True


def str_to_float(s: AnyStr):
    """Принимает str, "по-умному" возвращает float"""
    if not isinstance(s, str):
        raise TypeError(f'Ожидалось str, приняло {type(s)}')
    return float(s.replace(' ', '').replace(',', '.'))


def tests():
    assert prime_num_check(1) == False
    assert prime_num_check(2) == True
    assert prime_num_check(3) == True
    assert prime_num_check(4) == False
    assert prime_num_check(5) == True
    assert prime_num_check(6) == False
    assert prime_num_check(13) == True
    assert prime_num_check(4.0) == False
    assert prime_num_check(-3.7) == False
    assert prime_num_check(-7) == False
    assert prime_num_check('-6') == False
    assert prime_num_check('Test') == False
    assert prime_num_check(None) == False

    assert str_to_float('0 ,  888') == 0.888
    assert str_to_float('88') == 88
    assert str_to_float(',11') == 0.11
    assert str_to_float('. 12') == 0.12
    print('Тесты пройдены!\n')


if __name__ == '__main__':
    print('\nПривет! Это тестовое задание в компании RUVENTS!\n')
    tests()
    count_num1()
    count_num2()
    count_num3()
    count_date1()
    count_date2()
    count_date3()
