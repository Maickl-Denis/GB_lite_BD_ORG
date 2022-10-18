import staff_handler as sh
import client_handler as ch
import order_handler as oh
import os

def menu():
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')  
        num = first_menu()
        match num:
            # 1. Продажа
            case 1:
                os.system('cls' if os.name == 'nt' else 'clear')
                num1 = order_menu()
                match num1:
                    #1. Создать бланк заказа
                    case 1:
                        oh.create_order_list()
                    #2. Список заказов 
                    case 2:
                        oh.order_list()
                    case 0:
                        continue
                    case _:
                        print("Нет такого пункта меню")
            # 2. Работы с клиентами
            case 2:
                os.system('cls' if os.name == 'nt' else 'clear')
                num1 = clients_menu()
                match num1:
                    # 1. Список клиентов
                    case 1:
                        ch.client_list()
                    # 2. Добавить клиента
                    case 2:
                        ch.add_client()
                    case 0:
                        continue
                    case _:
                        print("Нет такого пункта меню")
            # 3. Работы с персоналом
            case 3:
                os.system('cls' if os.name == 'nt' else 'clear')
                num1 = person_menu()
                match num1:
                    # 1. Список сотрудников
                    case 1:
                        sh.staff_list()
                    # 2. Добавить сорудника
                    case 2:
                        sh.staff_replenishment()
                    # 3. Увольнение сотрудника
                    case 3:
                        sh.fired_replenishment(sh.staff_selection_id())
                    # 4. Справичник работ
                    case 4:
                        os.system('cls' if os.name == 'nt' else 'clear')
                        num1 = job_menu()
                        match num1:
                            # 1. Список работ
                            case 1:
                                sh.work_list()
                            # 2. Добавить работ
                            case 2:
                                sh.add_job()
                            # 3. Удаление работ
                            case 0:
                                continue
                    # 5. Добавить специализацию сотруднику
                    case 5:
                        sh.add_specialization(
                            sh.staff_selection_id(), sh.job_selection_id())
                    case 0:
                        continue
                    case _:
                        print("Нет такого пункта меню")
            case 0:
                break
            case _:
                print("Нет такого пункта меню")


def first_menu():
    try:
        num = int(input("1. Продажа\n"
                        "2. Работы с клиентами\n"
                        "3. Работы с персоналом\n"
                        "0. Выход\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        print("Вы ввели некоректное значение")
        return -1
    else:
        return num


def person_menu():
    try:
        num = int(input("1. Список сотрудников\n"
                        "2. Добавить сорудника\n"       
                        "3. Увольнение сотрудника\n"    
                        "4. Справичник работ\n"
                        "5. Добавить специализацию сотруднику\n"
                        "0. Выход\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        print("Вы ввели некоректное значение")
        return -1
    else:
        return num


def clients_menu():
    try:
        num = int(input("1. Список клиентов\n"
                        "2. Добавить клиента\n"
                        "0. Выход\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        print("Вы ввели некоректное значение")
        return -1
    else:
        return num


def job_menu():
    try:
        num = int(input("1. Список работ\n"
                        "2. Добавить вид работ\n"
                        "0. Выход\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        print("Вы ввели некоректное значение")
        return -1
    else:
        return num


def order_menu():
    try:
        num = int(input("1. Создать бланк заказа\n"
                        "2. Список заказов\n"
                        "0. Выход\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        print("Вы ввели некоректное значение")
        return -1
    else:
        return num