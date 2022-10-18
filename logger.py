from datetime import datetime
from time import time

def operation_logger(data):
    time = datetime.now().strftime('%d.%m.%y %H:%M')
    with open('log_barber.txt', 'a', encoding='utf-8') as file:
        file.write(f'{time} - {data}\n')


