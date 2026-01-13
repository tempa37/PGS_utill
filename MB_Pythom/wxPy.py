from tkinter import *
from tkinter import ttk
from tkinter import scrolledtext  
import Sh
from Sh import sendCmdBlock

# Параметры для файла Sh.py
scanIP = "192.168.70.33"
scanId = 5

tabFile = "PgsDiSetup.xlsx"
rowSt = 5
rowCnt = 11
colSt = 7
maxCmdInRow = 6
text_CmdBlk = ""

pos_left = 30
scan_pos_top = 10
pos_step = 26

tab_pos_top = 100
entry_cnt_size = 40

root = Tk()     # создаем корневой объект - окно

root.title("Конфигуратор ПГС для СКАН")     # устанавливаем заголовок окна
root.geometry("600x370+400+200")    # устанавливаем размеры окна
 
# Параметры подключения СКАН
lblScanParam = Label(text="Параметры подключения СКАН:", font=("System", 11))
lblScanParam.place(x=pos_left, y=scan_pos_top)

lblScanIP = Label(text="IP алрес:")
lblScanIP.place(x=pos_left, y=scan_pos_top + pos_step * 1)

lblScanID = Label(text="MODBUS ID:")
lblScanID.place(x=pos_left, y=scan_pos_top + pos_step * 2)

# Параметры таблицы команд
lblTableParam = Label(text="Параметры таблицы команд:", font=("System", 11))
lblTableParam.place(x=pos_left, y=tab_pos_top)

lblTableFile = Label(text="Имя файла команд (.xlsx):")
lblTableFile.place(x=pos_left, y=tab_pos_top + pos_step * 1)

lblTableHead = Label(text="Начальная   Количество")
lblTableHead.place(x=pos_left + 60, y=tab_pos_top + pos_step * 3)

lblTableRow = Label(text="Строки:")
lblTableRow.place(x=pos_left, y=tab_pos_top + pos_step * 4)

lblTableCol = Label(text="Колонки:")
lblTableCol.place(x=pos_left, y=tab_pos_top + pos_step * 5)

lblTableCol = Label(text="Максим. колич. команд в строке:")
lblTableCol.place(x=pos_left, y=tab_pos_top + pos_step * 6)

# Поля ввода
entryIP = ttk.Entry() # IP адрес 192.168.70.33
entryIP.place(x=pos_left + 100, y=scan_pos_top + pos_step * 1, width = 90)
entryIP.insert(0, scanIP)

entryID = ttk.Entry() # MODBUS ID 5
entryID.place(x=pos_left + 100, y=scan_pos_top + pos_step * 2, width = entry_cnt_size)
entryID.insert(0, str(scanId))

entryFile = ttk.Entry() # Файл таблицы команд
entryFile.place(x=pos_left + 50, y=tab_pos_top + pos_step * 2, width = 150)
entryFile.insert(0, tabFile)

entryRowSt = ttk.Entry() # Строки нач
entryRowSt.place(x=pos_left + 70, y=tab_pos_top + pos_step * 4, width = entry_cnt_size)
entryRowSt.insert(0, str(rowSt))

entryRowCnt = ttk.Entry() # Строки кол 
entryRowCnt.place(x=pos_left + 70 + 70, y=tab_pos_top + pos_step * 4, width = entry_cnt_size)
entryRowCnt.insert(0, str(rowCnt))

entryColSt = ttk.Entry() # Колонки нач
entryColSt.place(x=pos_left + 70, y=tab_pos_top + pos_step * 5, width = entry_cnt_size)
entryColSt.insert(0, str(colSt))

entryCmdCnt = ttk.Entry() # Количество команд в строке
entryCmdCnt.place(x=pos_left + 70, y=tab_pos_top + pos_step * 7, width = entry_cnt_size)
entryCmdCnt.insert(0, str(maxCmdInRow))

textResult = scrolledtext.ScrolledText(root, width=35, wrap="word", bg="black", fg="white")  
textResult.pack(anchor=E, fill=Y)

# for i in range(1, 25):
    # textResult.insert(END, "Текст на черном\n")

# btn = Button(text="Пуск") # создаем кнопку из пакета tkinter

cmdText_Val = "" #"Команды настройки:\n"
def cmdUpdate(updText):
    textResult.insert(END, updText)
    textResult.insert(END, "\n")

def click_Pusk():
    textResult.delete("1.0", END)
    textResult.insert(END, "Передача команд настройки\n\n")
    root.update()
    tabFile = entryFile.get()
    rowSt = int(entryRowSt.get())
    rowCnt = int(entryRowCnt.get())
    colSt = int(entryColSt.get())
    maxCmdInRow = int(entryCmdCnt.get())
    
    cmdText_Val = sendCmdBlock(tabFile, rowSt, rowCnt, colSt, maxCmdInRow)
    textResult.insert(END, cmdText_Val)
    # print(cmdText_Val)
    root.update()
    
 
def delete_text():
    textResult.delete("1.0", END)
    
def click_Exit():
    root.destroy()  # ручное закрытие окна и всего приложения
    print("Конец сеанса")

# btnPusk = ttk.Button(text="Click Me", state=["disabled"])
btnPusk = ttk.Button(text="Пуск", command=click_Pusk)
btnPusk.place(x=50, y=320)

btnExit = ttk.Button(text="Выход", command=click_Exit)
btnExit.place(x=170, y=320)

icon = PhotoImage(file = "Save32.png")
root.iconphoto(False, icon)

root.update_idletasks()
# print(root.geometry())    # "300x250+400+200"

def print_info(widget, depth=0):
    widget_class=widget.winfo_class()
    widget_width = widget.winfo_width()
    widget_height = widget.winfo_height()
    widget_x = widget.winfo_x()
    widget_y = widget.winfo_y()
    print("   "*depth + f"{widget_class} width={widget_width} height={widget_height}  x={widget_x} y={widget_y}")
    for child in widget.winfo_children():
        print_info(child, depth+1)
 
root.update()     # обновляем информацию о виджетах
 
# print_info(root)

root.protocol("WM_DELETE_WINDOW", click_Exit)

root.mainloop()

