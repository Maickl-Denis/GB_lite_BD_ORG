import os
import menu
import creat_bd

if __name__ == '__main__':

    if  not os.path.exists("base.xlsx"):
        creat_bd.cr_db()
    menu.menu()