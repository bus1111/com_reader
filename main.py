import os
import sys
import serial
from serial.tools import list_ports
from datetime import datetime
from threading import Thread
from openpyxl import Workbook


PREFIX = b'sendordata'


def create_workbook(sensor_count):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Сенсоры'
    ws.cell(1, 1, 'Время')
    for i in range(1, sensor_count + 1):
        ws.cell(row=1, column=i + 1, value=f'Сенсор {i}')
    return wb


def reader(wb, port, sensor_count, should_stop):
    ws = wb.active
    with serial.Serial(port.device, 115200, timeout=0.25) as ser:
        while True:
            try:
                line = ser.readline()
                if should_stop():
                    print()
                    break
                if len(line) == 0 or not line.startswith(PREFIX):
                    continue
            except Exception as e:
                print('\nОшибка чтения, запись закончена:', e)
                break

            try:
                line = line.strip().replace(PREFIX, b'').decode('utf-8')
                values = line.split('/')
                print(f'\rПоследние значения: {", ".join(values):{sensor_count*8}}', end='')
                values = [int(v) for v in values[:sensor_count]]
                ws.append([datetime.now()] + values)
            except Exception as e:
                print('\rОшибка обработки:', e, end='')
            sys.stdout.flush()


def menu():
    while True:
        ports = sorted(list_ports.comports())
        if len(ports) == 0:
            print('Устройства не найдены, повторите.')
            input()
            continue
        break

    print('Доступные устройства:')
    for i, port in enumerate(ports):
        print(f'   * {i + 1}: {port.name} - {port.description}')

    while True:
        try:
            selected_port = 1
            selected_port = input(f'Выберите устройство [{selected_port}]: ') or selected_port
            selected_port = int(selected_port)
            if selected_port > len(ports) or selected_port <= 0:
                print('Неверный ввод')
            else:
                break
        except Exception:
            print('Неверный ввод')
            continue

    while True:
        try:
            sensor_count = 3
            sensor_count = input(f'Введите количество сенсоров [{sensor_count}]: ') or sensor_count
            sensor_count = int(sensor_count)
            break
        except Exception:
            print('Неверный ввод')
            continue

    return ports[selected_port - 1], sensor_count


if __name__ == '__main__':
    while True:
        port, sensor_count = menu()
        wb = create_workbook(sensor_count)
        stop_signal = False
        t = Thread(target=reader, args=(wb, port, sensor_count, lambda: stop_signal))
        t.start()
        input('Запись начата, нажмите Enter чтобы остановить\n')
        stop_signal = True
        t.join(timeout=5)
        filename = f'{datetime.strftime(datetime.now(), "Сенсоры %d %b %Y %H-%M-%S.xlsx")}'
        wb.save(filename)
        print(f'Запись сохранена в файл {os.path.abspath(filename)}')
        input('Нажмите Enter чтобы начать новую запись')
