"""
Модуль обработки заказов

"""


import openpyxl
from openpyxl import load_workbook
import staff_handler as sh
import client_handler as ch
from logger import operation_logger as logg


def create_order_list():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 4
    i = 2
    while sheet.active[f"A{i}"].value:
        i += 1
    sheet.active[f"A{i}"].value = i - 1
    while True:
        id_job = sh.job_selection_id()
        sheet.active[f"B{i}"].value = id_job
        exit_key = sh.staff_selection_id(True, id_job)
        if exit_key == -1:
            break
        elif exit_key == 0:
            continue
        else:
            id_client = ch.client_selection_id()
            sheet.active[f"D{i}"].value = exit_key
            sheet.active[f"C{i}"].value = price_job(id_job)
            sheet.active[f"E{i}"].value = id_client
            sheet.save('base.xlsx')
            break
    sheet.active = 2
    data_job = f'{sheet.active[f"B{id_job+1}"].value}'
    sheet.active = 0
    data_staff = f'{sheet.active[f"B{exit_key+1}"].value}{sheet.active[f"C{exit_key+1}"].value}'
    sheet.active = 3
    data_client = f'{sheet.active[f"B{id_client+1}"].value} {sheet.active[f"C{id_client+1}"].value}'
    logg(f"Оказана услуга {data_job} клиенту {data_client} мастером {data_staff}цена: {price_job(id_job)}")
    input("Для продолжения работы нажмите Enter...")


def price_job(id_job: int):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 2
    return sheet.active[f"C{id_job+1}"].value

#!Список заказов


def order_list():
    sheet = openpyxl.open("base.xlsx")
    print("\033[0m"+("_" * 160))
    print(
        f'     {"Номер заказа":28}{"Название работы":30}{"Цена":30}{"Мастер":40}{"Клиент"}')
    i = 2
    while sheet.active[f"A{i}"].value:
        sheet.active = 4
        id_job = sheet.active[f"B{i}"].value
        id_staff = sheet.active[f"D{i}"].value
        id_client = sheet.active[f"E{i}"].value
        sheet.active = 2
        data_job = f'{sheet.active[f"B{id_job+1}"].value}'
        sheet.active = 0
        data_staff = f'{sheet.active[f"B{id_staff+1}"].value} {sheet.active[f"C{id_staff+1}"].value}'
        sheet.active = 3
        data_client = f'{sheet.active[f"B{id_client+1}"].value} {sheet.active[f"C{id_client+1}"].value}'
        sheet.active = 4
        print(f'{"|"}{sheet.active[f"A{i}"].value:^22} {"|":12} {data_job:<17} {"|":<7} {sheet.active[f"C{i}"].value:<10} {"|":<14} {data_staff:<25} {"|":<10} {data_client:<31} {"|":<10}')
        i += 1
    print("\033[0m"+("_" * 160))
    logg("Выведен список заказов")
    input("Для продолжения работы нажмите Enter...")
