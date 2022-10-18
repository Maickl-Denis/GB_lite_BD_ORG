import os
from menu import menu
import creat_bd
from logger import operation_logger as logg

if __name__ == '__main__':

    if  not os.path.exists("base.xlsx"):
        logg("Не наедена база данных, создается новая")
        creat_bd.cr_db()
    logg("Запуск программы")
    menu()