from tkinter import *
from tkinter import ttk
from tkinter import scrolledtext  
import Sh
from Sh import sendCmdBlock, export_posts_to_excel

# Параметры для файла Sh.py
scanIP = "192.168.70.33"
scanId = 5

tabFile = "PgsDiSetup.xlsx"
rowSt = 5
rowCnt = 11
colSt = 7
maxCmdInRow = 6
text_CmdBlk = ""

# === Добавлено: параметры экспорта ===
export_post_count = 10
export_filename = "registers_pgs"
# === Конец добавления: параметры экспорта ===

pos_left = 30
scan_pos_top = 10
pos_step = 26

tab_pos_top = 100
entry_cnt_size = 40

root = Tk()     # создаем корневой объект - окно

root.title("Конфигуратор ПГС для СКАН")     # устанавливаем заголовок окна
root.geometry("660x370+400+200")    # устанавливаем размеры окна
root.minsize(660, 370)
 
 
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


# === Добавлено: диалог экспорта в Excel ===
def open_export_dialog():
    dialog = Toplevel(root)
    dialog.title("Экспорт")
    dialog.geometry("360x190+450+250")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()

    count_var = StringVar(value="10")
    filename_var = StringVar(value="registers_pgs")

    def validate_count(value):
        return value.isdigit() or value == ""

    def save_export_settings():
        global export_post_count, export_filename

        count_value = count_var.get()
        if not count_value.isdigit():
            count_value = "10"
        count_int = int(count_value)
        if count_int > 50:
            count_int = 50
        if count_int < 1:
            count_int = 1

        export_post_count = count_int
        export_filename = filename_var.get().strip() or "registers_pgs"

        scan_ip_value = entryIP.get().strip() or scanIP
        scan_id_value = entryID.get().strip()
        if scan_id_value.isdigit():
            scan_id_value = int(scan_id_value)
        else:
            scan_id_value = scanId

        textResult.insert(END, "Экспорт данных...\n")
        root.update()
        exported_file = export_posts_to_excel(
            export_post_count,
            export_filename,
            scan_ip_value,
            scan_id_value,
            output_callback=cmdUpdate,
        )
        textResult.insert(END, f"Экспорт завершен: {exported_file}\n")
        root.update()

        dialog.destroy()

    def close_dialog():
        dialog.destroy()

    lblCount = Label(dialog, text="Количество постов:")
    lblCount.place(x=20, y=20)

    count_validate = (dialog.register(validate_count), "%P")
    entryCount = ttk.Entry(dialog, textvariable=count_var, validate="key", validatecommand=count_validate)
    entryCount.place(x=190, y=20, width=60)

    lblFilename = Label(dialog, text="Название файла:")
    lblFilename.place(x=20, y=60)

    entryFilename = ttk.Entry(dialog, textvariable=filename_var)
    entryFilename.place(x=190, y=60, width=120)

    lblExtension = Label(dialog, text=".xlsx")
    lblExtension.place(x=315, y=60)

    btnExportAction = ttk.Button(dialog, text="Экспортировать", command=save_export_settings)
    btnExportAction.place(x=200, y=120, width=120)

    btnBack = ttk.Button(dialog, text="Назад", command=close_dialog)
    btnBack.place(x=20, y=120, width=140)

    entryCount.focus()
# === Конец добавления: диалог экспорта в Excel ===

# btnPusk = ttk.Button(text="Click Me", state=["disabled"])
btnPusk = ttk.Button(text="Пуск", command=click_Pusk)
btnPusk.place(x=50, y=320)

# === Добавлено: кнопка экспорта ===
btnExport = ttk.Button(text="Экспорт", command=open_export_dialog)
btnExport.place(x=140, y=320)
# === Конец добавления: кнопка экспорта ===

btnExit = ttk.Button(text="Выход", command=click_Exit)
btnExit.place(x=250, y=320)

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
