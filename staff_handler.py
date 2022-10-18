"""
Модуль работы с персоналом

"""



import os
import openpyxl
from openpyxl import load_workbook
from logger import operation_logger as logg


#!добавить персонал
def staff_replenishment():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 0
    i = 2
    while sheet.active[f"A{i}"].value:
        i += 1
    name = input("\033[1m" + "Введите имя сотрудника: ")
    name = "-" if len(name) == 0 else name
    while name.isdigit():
        name = input(f'\033[31m{"Ошибка!"}\n'
                         f'\033[0m\033[1m{"Введите имя сотрудника: "}')
        name = "-" if len(name) == 0 else name
    sec_name = input("\033[1m" + "Введите фамилию сотрудника: ")
    sec_name = "-" if len(sec_name) == 0 else sec_name
    while sec_name.isdigit():
        sec_name = input(f'\033[31m{"Ошибка!"}\n'
                      f'\033[0m\033[1m{"Введите фамилию сотрудника: "}')
        sec_name = "-" if len(sec_name) == 0 else sec_name
    phone = input("\033[1m" + "Введите номер телефона сотрудника: ")
    while not phone.isdigit():
        phone = input(f'\033[31m{"Ошибка!"}\n'
                      f'\033[0m\033[1m{"Введите номер телефона сотрудника: "}')
    spec = input("\033[1m" + "Введите описание: ")
    spec = "-" if len(spec) == 0 else spec
    sheet.active[f"A{i}"].value = i-1
    sheet.active[f"B{i}"].value = name
    sheet.active[f"C{i}"].value = sec_name
    sheet.active[f"D{i}"].value = phone
    sheet.active[f"E{i}"].value = spec
    sheet.active[f"F{i}"].value = 1
    sheet.save('base.xlsx')
    logg(f"Добавлен сотрудник: {name} {sec_name}, тел: {phone}. Описание: {spec}")


#!добавить работу
def add_job():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 2
    i = 2
    print("\033[0m" + ("_" * 160))
    while sheet.active[f"A{i}"].value:
        i += 1
    job_name = input("\033[1m" + "Введите название работы: ")
    job_name = "-" if len(job_name) == 0 else job_name
    while job_name.isdigit():
        job_name = input(f'\033[31m{"Ошибка!"}\n'
                   f'\033[0m\033[1m{"Введите название работы: "}')
        job_name = "-" if len(job_name) == 0 else job_name
    price = input("\033[1m" + "Введите стоимость: ")
    while not price.isdigit():
        price = input(f'\033[31m{"Ошибка!"}\n'
                         f'\033[0m\033[1m{"Введите стоимость: "}')

    spec = input("\033[1m" + "Введите описание: ")
    spec = "-" if len(spec) == 0 else spec
    sheet.active[f"A{i}"].value = i-1
    sheet.active[f"B{i}"].value = job_name
    sheet.active[f"C{i}"].value = price
    sheet.active[f"D{i}"].value = spec
    sheet.save('base.xlsx')
    logg(f"Добавлена работа: {job_name} Цена: {price}. Описание: {spec}")


#!добавить специализацию
def add_specialization(id_staff: int, id_job: int):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 1
    i = 2
    while sheet.active[f"A{i}"].value:
        i += 1
    sheet.active[f"A{i}"].value = id_staff
    sheet.active[f"B{i}"].value = id_job
    sheet.save('base.xlsx')
    sheet.active = 0
    lg1 = f'{sheet.active[f"B{id_staff+1}"].value}{sheet.active[f"C{id_staff+1}"].value}'
    sheet.active = 2
    lg2 = f'{sheet.active[f"B{id_job+1}"].value}'
    logg(f"Сотруднику: {lg1} добавлена специализация {lg2}")


#!выбор персонал
def staff_selection_id(check_specialization_trigger=False, id_job=None):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 0
    id_master = {}
    i = 2
    staff_number = 1
    print("\033[0m"+("_" * 160))
    print("Пресонал: ")
    if check_specialization_trigger == True:
        while sheet.active[f"A{i}"].value:
            if sheet.active[f"F{i}"].value and check_specialization(int(sheet.active[f"A{i}"].value), id_job):
                id_master.update({staff_number: sheet.active[f"A{i}"].value})
                print(
                    f'  {staff_number}.{sheet.active[f"B{i}"].value} {sheet.active[f"C{i}"].value}')
                staff_number += 1
            i += 1

        if not len(id_master):
            print("\033[31m"+"   Нет подходящего мастера!\n")

            print("\033[0m"+"Извините,эта услуга временно недоступна:(\n"
                  "Хотите выбрать другую?\n"
                  "1. да\n"
                  "2. нет")
            print("\033[0m" + ("_" * 160))
            client_choice = input("\033[1m"+"Ввод: ")
            while not client_choice.isdigit() or int(client_choice) not in [1, 2]:
                client_choice = input(f'\033[31m{"Ошибка!"}\n'
                                      f'\033[0m\033[1m{"Ввод: "}')
            if int(client_choice) == 1:
                return 0
            else:
                return -1
    else:
        while sheet.active[f"A{i}"].value:
            if sheet.active[f"F{i}"].value:
                id_master.update({staff_number: sheet.active[f"A{i}"].value})
                print(
                    f'  {staff_number}.{sheet.active[f"B{i}"].value} {sheet.active[f"C{i}"].value}')
                staff_number += 1
            i += 1
            
    print("\033[0m" + ("_" * 160))
    id = input('\033[1m'+'Введите цифру нужного мастера: ')
    while not id.isdigit() or int(id) not in id_master.keys():
        id = input(f'\033[31m{"Ошибка!"}\n'
                   f'\033[0m\033[1m{"Введите цифру нужного мастера: "}')
    return int(id_master[int(id)])



#!выбор работ 
def job_selection_id():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 2
    i = 2
    print("\033[0m" + ("_" * 160))
    print("Виды работ: ")
    while sheet.active[f"A{i}"].value:
        print(
            f'  {sheet.active[f"A{i}"].value}.{sheet.active[f"B{i}"].value} {sheet.active[f"C{i}"].value}')
        i += 1
    print("\033[0m" + ("_" * 160))
    id = input(f'\033[1m{"Введите цифру нужной работы: "}')
    while not id.isdigit() or int(id) > i - 2 or int(id) <= 0:
        id = input(f'\033[31m{"Ошибка!"}\n'
                   f'\033[0m\033[1m{"Введите цифру нужной работы: "}')
    return int(id)

#!Список персонала
def staff_list():
    os.system('cls' if os.name == 'nt' else 'clear')
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 0
    i = 2
    print("\033[0m" + ("_" * 160))
    print(
        f'     {"Номер сотрудника":28}{"Имя":21}{"Фамилия":25}{"Телефон":33}{"Описание":33}{"Статус"}')
    print("\033[0m" + ("-" * 160))
    while sheet.active[f"A{i}"].value:
        status = "Работает" if sheet.active[f"F{i}"].value else "Уволен"
        print(f'{"|":12}{sheet.active[f"A{i}"].value:<14}{"|":6}{sheet.active[f"B{i}"].value:<12}{"|":10}{sheet.active[f"C{i}"].value:<15}{"|":6}{sheet.active[f"D{i}"].value:<20}{"|":4}{sheet.active[f"E{i}"].value:^36}{"|":9}{status:<15}{"|"}')
        i += 1
    print("\033[0m" + ("_" * 160))
    logg("Выведен список персонала")
    input("Для продолжения работы нажмите Enter...")

#!Увольнение сотрудника
def fired_replenishment(id_staff:int):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 0
    sheet.active[f"F{id_staff+1}"].value = 0
    sheet.save('base.xlsx')
    staf = f'{sheet.active[f"B{id_staff+1}"].value}{sheet.active[f"C{id_staff+1}"].value}'
    print("\033[0m" + ("_" * 160))
    print("\033[31m"+f"Сотрудник {staf} уволен(a)!"+"\033[0m")
    logg(f"Сотруднику: {staf} уволен(a)")
    input("Для продолжения работы нажмите Enter...")


def check_specialization(id_staff: int, id_job: int):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 1
    i = 2
    while sheet.active[f"A{i}"].value:
        if sheet.active[f"A{i}"].value == id_staff and sheet.active[f"B{i}"].value == id_job:
            return True
        i += 1
    return False

#!Список работ
def work_list():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 2
    i = 2
    print("\033[0m" + ("_" * 160))
    print(
        f'     {"Номрер работы":30}{"Название работы":37}{"Цена":50}{"Описание"}')
    print("\033[0m" + ("-" * 160))
    while sheet.active[f"A{i}"].value:
        status = "Работает" if sheet.active[f"F{i}"].value else "Уволен"
        print(
            f'{"|":12}{sheet.active[f"A{i}"].value:<14}{"|":6}{sheet.active[f"B{i}"].value:<30}{"|":10}{sheet.active[f"C{i}"].value:<15}{"|":7}{sheet.active[f"D{i}"].value:^65}{"|"}')
        i += 1
    print("\033[0m" + ("-" * 160))
    logg("Выведен список работ")
    input("Для продолжения работы нажмите Enter...")