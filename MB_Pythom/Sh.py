from tkinter import *
from tkinter import ttk
from tkinter import scrolledtext  

import sys

import os
import time
# import shutil
from openpyxl import load_workbook

import pyModbusTCP
from pyModbusTCP.client import ModbusClient

# global scanIP, scanId, tabFile, rowSt, rowCnt, colSt, maxCmdInRow
# import wxPy
# from wxPy import cmdUpdate

def sendCmdBlock(tabFile, rowSt, rowCnt, colSt, maxCmdInRow):
    #Имя файла с командами для настройки ПГС
    file_path = 'PgsSetup.xlsx'
    # позиция и размер таблицы
    start_row = rowSt
    row_count = 3
    start_col = 3
    col_count = 12

    # if __name__ == '__main__':
    cmdText_l = "Команды настройки:\n"

    cells_in_cmd = 3
    scan_IP = '192.168.70.33'
    scan_ID = 5
    if len(sys.argv) > 5:
        file_path = sys.argv[1]
        start_row = int(sys.argv[2])
        row_count = int(sys.argv[3])
        start_col = int(sys.argv[4])
        col_count = int(sys.argv[5]) * cells_in_cmd
        if len(sys.argv) > 6:
            scan_IP = sys.argv[6]
        if len(sys.argv) > 7:
            scan_ID = int(sys.argv[7])

        # print(f'start_row = {start_row}, row_count = {row_count}')
        # print(f'start_col = {start_col}, col_count = {col_count}')
    else:
        file_path = tabFile
        start_row = int(rowSt)
        row_count = int(rowCnt)
        start_col = int(colSt)
        col_count = int(maxCmdInRow) * cells_in_cmd
    
    last_row = start_row + row_count - 1
    err = True
    # Проверяем, существует ли файл
    if os.path.exists(file_path):
        # print(f"Есть Файл {file_path}")
        workbook = load_workbook(file_path)
        page = workbook.active
        
        client = ModbusClient(host = scan_IP, port = 502, unit_id = scan_ID, timeout = 0.5, auto_open = True)
    #    ModbusClient(host="192.168.56.101", port=502, debug= False, unit_id=5, auto_open=True, auto_close=True)
        curr_row = start_row
        for row in page.iter_rows(min_row = start_row, max_row = last_row, values_only=True):
            for i in range(start_col, start_col + col_count, cells_in_cmd):
                curr_value = page.cell(row = curr_row, column = i + 1).value
                # print(curr_row, i, curr_value)
                if (curr_value == None):
                    break
                else:
                    aVal = [row[i + 0], row[i + 1], row[i + 2]]
                    
                    aCmd = [row[i + 0], row[i + 1], row[i + 2]]
                    for i in range(0, 2):
                        if (aVal[i] < 0):
                            aCmd[i] += 65536
                    err = client.write_multiple_registers(2901, aCmd)
                    if (err == False):
                        # print(f"{aVal} Нет ответа!")
                        cmd = f"{aVal} Нет ответа!"
                    else:
                        # print(aVal)
                        cmd = f"{aVal}"
                    # print(cmd)
                    cmdText_l += (cmd + "\n")
                    if (err == False):
                        break
            curr_row = curr_row + 1
            if (err == False):
                break
        client.close()
    else:
        print(f"Нет Файла {file_path}")
    return cmdText_l
