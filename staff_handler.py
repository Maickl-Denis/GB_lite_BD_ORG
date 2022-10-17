import openpyxl
from openpyxl import load_workbook


#+добавить персонал
def staff_replenishment(name_s: str,sec_name: str,phone: str,spec:str):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 0
    i = 2
    while sheet.active[f"A{i}"].value:
        i += 1
    sheet.active[f"A{i}"].value = i-1
    sheet.active[f"B{i}"].value = name_s
    sheet.active[f"C{i}"].value = sec_name
    sheet.active[f"D{i}"].value = phone
    sheet.active[f"E{i}"].value = spec
    sheet.active[f"F{i}"].value = 1
    sheet.save('base.xlsx')


#+добавить работу
def add_job(job_id: str,price: str,spec:str):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 2
    i = 2
    while sheet.active[f"A{i}"].value:
        i += 1
    sheet.active[f"A{i}"].value = i-1
    sheet.active[f"B{i}"].value = job_id
    sheet.active[f"C{i}"].value = price
    sheet.active[f"D{i}"].value = spec
    sheet.save('base.xlsx')


#+добавить специализацию
def add_specialization(id_sotr: str,vid_rab: str):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 1
    i = 2
    while sheet.active[f"A{i}"].value:
        i += 1
    sheet.active[f"A{i}"].value = id_sotr
    sheet.active[f"B{i}"].value = vid_rab
    sheet.save('base.xlsx')


#+выбор персонал
def staff_selection_id(spec_trigger = False,job_id = None):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 0
    i = 2
    print("Пресонал: ")
    if spec_trigger == True:
        while sheet.active[f"A{i}"].value:
            if sheet.active[f"F{i}"].value and check_specialization(job_id,int(sheet.active[f"A{i}"].value)):
                print(f'  {sheet.active[f"A{i}"].value}.{sheet.active[f"B{i}"].value} {sheet.active[f"C{i}"].value}')
            i+=1
    else:
        while sheet.active[f"A{i}"].value:
            if sheet.active[f"F{i}"].value:
                print(f'  {sheet.active[f"A{i}"].value}.{sheet.active[f"B{i}"].value} {sheet.active[f"C{i}"].value}')
            i+=1
    id = int(input(f'Введите цифру нужного мастера: '))
    while id > i-2 or id <= 0:
        print("Ошибка такого мастера не существует!")
        id = int(input(f'Введите цифру нужного мастера: '))
    return id



#+выбор работ 
def job_selection_id():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 2
    i = 2
    print("Виды работ: ")
    while sheet.active[f"A{i}"].value:
        print(f'  {sheet.active[f"A{i}"].value}.{sheet.active[f"B{i}"].value} {sheet.active[f"C{i}"].value}')
        i+=1
    id = int(input(f'Введите цифру нужной работы: '))
    while id > i-2 or id <= 0:
        print("Ошибка!")
        id = int(input(f'Введите цифру нужной работы: '))
    return id

#+Список персонала
def staff_list():
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 0
    i = 2
    while sheet.active[f"A{i}"].value:
        if sheet.active[f"F{i}"].value:
            print(f'  {sheet.active[f"A{i}"].value}.{sheet.active[f"B{i}"].value} {sheet.active[f"C{i}"].value} {sheet.active[f"D{i}"].value} {sheet.active[f"E{i}"].value}')
        i+=1
        
#+Увольнение сотрудника
def fired_replenishment(id_sotr:int):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 0
    sheet.active[f"F{id_sotr+1}"].value = 0
    sheet.save('base.xlsx')


def check_specialization(job_id,id_sotr):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 1
    i = 2
    while sheet.active[f"A{i}"].value:
        if sheet.active[f"A{i}"].value == id_sotr and sheet.active[f"B{i}"].value == job_id:
            return True
        i+=1
    return False