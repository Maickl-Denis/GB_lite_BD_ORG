"""
Модуль работы с клиентами

"""
import openpyxl
from openpyxl import load_workbook
from logger import operation_logger as logg


#!добавить клиента
def add_client():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 3
    i = 2
    while sheet.active[f"A{i}"].value:
        i += 1
    print("\033[0m" + ("_" * 160))
    name = input("\033[1m" + "Введите имя клиента: ")
    while len(name) == 0 or name.isdigit():
        name = input(f'\033[31m{"Ошибка!"}\n'
                     f'\033[0m\033[1m{"Введите имя клиента: "}')
    sec_name = input("\033[1m" + "Введите фамилию клиента: ")
    while len(sec_name) == 0 or sec_name.isdigit():
        sec_name = input(f'\033[31m{"Ошибка!"}\n'
                         f'\033[0m\033[1m{"Введите фамилию клиента: "}')
    phone = input("\033[1m" + "Введите номер телефона клиента: ")
    while not phone.isdigit():
        phone = input(f'\033[31m{"Ошибка!"}\n'
                      f'\033[0m\033[1m{"Введите номер телефона клиента: "}')
    spec = input("\033[1m" + "Введите описание: "+"\033[0m")
    while len(spec) == 0:
        spec = input(f'\033[31m{"Ошибка!"}\n'
                     f'\033[0m\033[1m{"Введите описание: "}\033[0m')
    sheet.active[f"A{i}"].value = i-1
    sheet.active[f"B{i}"].value = name
    sheet.active[f"C{i}"].value = sec_name
    sheet.active[f"D{i}"].value = phone
    sheet.active[f"E{i}"].value = spec
    sheet.save('base.xlsx')
    logg(f"Добавлен клиент: {name} {sec_name}, тел: {phone}. Описание: {spec}")


#!список всех клиентов
def client_list():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 3
    i = 2
    print("\033[0m" + ("_" * 136))
    print(
        f'     {"Номер клиента":28}{"Имя":21}{"Фамилия":25}{"Телефон":33}{"Описание"}')
    print("\033[0m" + ("-" * 136))
    while sheet.active[f"A{i}"].value:
        print(f'{"|":12}{sheet.active[f"A{i}"].value:<14}{"|":6}{sheet.active[f"B{i}"].value:<12}{"|":10}{sheet.active[f"C{i}"].value:<15}{"|":6}{sheet.active[f"D{i}"].value:<20}{"|":4}{sheet.active[f"E{i}"].value:^36}{"|"}')
        i += 1
    print("\033[0m" + ("-" * 136))
    logg("Выведен список клиентов")
    input("Для продолжения работы нажмите Enter...")



def client_selection_id():
    while True:
        sheet = openpyxl.open("base.xlsx")
        sheet.active = 3
        i = 2
        print("\033[0m" + ("_" * 160))
        print("Выберите клиента: ")
        while sheet.active[f"A{i}"].value:
            print(
                f'  {sheet.active[f"A{i}"].value}.{sheet.active[f"B{i}"].value} {sheet.active[f"C{i}"].value} ')
            i += 1
        print("\033[31m" + "  0.Добавить клиента")
        print("\033[0m" + ("_" * 160))
        id = input(f'\033[1m{"Введите цифру клиента: "}\033[0m')
        if int(id) == 0:
            print("\033[0m" + ("_" * 160) + "\033[1m")
            add_client()
            continue
        while not id.isdigit() or int(id) > i - 2 or int(id) < 0:
            id = input(f'\033[31m{"Ошибка!"}\n'
                       f'\033[0m\033[1m{"Введите цифру клиента: "}')
        return int(id)