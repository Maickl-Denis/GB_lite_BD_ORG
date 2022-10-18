import staff_handler as sh
import client_handler as ch
import order_handler as oh
import os
from logger import operation_logger as logg

def menu():
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')  
        num, men = first_menu()
        logg(f"В главном меню: {'выбрано раздел' if men else 'введены не коректные данные'} {men if men else ''}")
        match num:
            # 1. Продажа
            case 1:
                os.system('cls' if os.name == 'nt' else 'clear')
                num1, men = order_menu()
                logg(f"В меню Продажа: {'выбрано раздел' if men else 'введены не коректные данные'} {men if men else ''}")
                match num1:
                    #1. Создать бланк заказа
                    case 1:
                        oh.create_order_list()
                    #2. Список заказов 
                    case 2:
                        oh.order_list()
                    case 0:
                        continue

            # 2. Работы с клиентами
            case 2:
                os.system('cls' if os.name == 'nt' else 'clear')
                num1, men = clients_menu()
                logg(f"В меню Работа с клиентами: {'выбрано раздел' if men else 'введены не коректные данные'} {men if men else ''}")
                match num1:
                    # 1. Список клиентов
                    case 1:
                        ch.client_list()
                    # 2. Добавить клиента
                    case 2:
                        ch.add_client()
                    case 0:
                        continue

            # 3. Работы с персоналом
            case 3:
                os.system('cls' if os.name == 'nt' else 'clear')
                num1, men = person_menu()
                logg(f"В меню Работа с персоналом: {'выбрано раздел' if men else 'введены не коректные данные'} {men if men else ''}")
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
                        num1, men = job_menu()
                        logg(f"В меню Справочник работ: {'выбрано раздел' if men else 'введены не коректные данные'} {men if men else ''}")
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

            case 0:
                logg("Программа завершила свою работу")
                break


def first_menu():
    try:
        ls = {1:"Продажа", 2: "Работа с клиентами", 3: "Работа с персоналом", 0: "Выход"}
        num = int(input(f"1. {ls[1]}\n"
                        f"2. {ls[2]}\n"
                        f"3. {ls[3]}\n"
                        f"0. {ls[0]}\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        return -1, None
    else:
        return num, ls.get(num)


def person_menu():
    try:
        ls = {1:"Список сотрудников", 2: "Добавить сорудника", 3: "Увольнение сотрудника",
              4: "Справичник работ", 5: "Добавить специализацию сотруднику", 0: "Выход"}
        num = int(input(f"1. {ls[1]}\n"
                        f"2. {ls[2]}\n"
                        f"3. {ls[3]}\n"
                        f"4. {ls[4]}\n"
                        f"5. {ls[5]}\n"
                        f"0. {ls[0]}\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        return -1, None
    else:
        return num, ls.get(num)


def clients_menu():
    try:
        ls = {1: "Список клиентов", 2: "Добавить клиента", 0: "Выход"}
        num = int(input(f"1. {ls[1]}\n"
                        f"2. {ls[2]}\n"
                        f"0. {ls[0]}\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        return -1, None
    else:
        return num, ls.get(num)


def job_menu():
    try:
        ls = {1: "Список работ", 2: "Добавить вид работ", 0: "Выход"}
        num = int(input(f"1. {ls[1]}\n"
                        f"2. {ls[2]}\n"
                        f"0. {ls[0]}\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        return -1, None
    else:
        return num, ls[num]


def order_menu():
    try:
        ls = {1: "Создать бланк заказа", 2: "Список заказов", 0: "Выход"}
        num = int(input(f"1. {ls[1]}\n"
                        f"2. {ls[2]}\n"
                        f"0. {ls[0]}\n"
                        "Выберете пункт меню:  "))
    except ValueError:
        return -1, None
    else:
        return num, ls[num]