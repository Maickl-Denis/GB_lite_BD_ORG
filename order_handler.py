import openpyxl
from openpyxl import load_workbook
import staff_handler as sh


def order_list(job_id: int,id_client: int):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 4
    i = 2
    while sheet.active[f"A{i}"].value:
        i += 1
    sheet.active[f"A{i}"].value = i-1
    sheet.active[f"B{i}"].value = job_id
    sheet.active[f"C{i}"].value = price_job(job_id)
    sheet.active[f"D{i}"].value = sh.staff_selection_id(True,job_id)
    sheet.active[f"E{i}"].value = id_client
    sheet.save('base.xlsx')

def price_job(job_id:int):
    sheet = openpyxl.open("base.xlsx")
    sheet.active = 2
    return int(sheet.active[f"C{job_id+1}"].value)