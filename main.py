from tkinter import messagebox, filedialog, Entry, Tk, StringVar, Label, END, Button
import re
from datetime import datetime
from helpers import processKinds

def selectFile():
    filename = filedialog.askopenfilename(
        initialdir = "/home/den/src/putevka/", title = "Выберитей файл путёвок!",
        filetypes = (("exel files","*.xlsx"), ("all files","*.*")))
    input_file_l['text'] = re.split(r'\\|/', filename)[-1]

def generate():
    #smena
    smena_no = smena_no_e.get()
    if len(smena_no) < 1:
        messagebox.showwarning("Номер смены","Заполните номер смены!")
        return
    #date
    smena_date = smena_date_e.get()
    try:
        d,m,y = smena_date.split('.')
    except:
        messagebox.showwarning("Заполните дату!","Неправильно заполнена дата!")
        return
    try:
        smena_date = datetime(int(y), int(m), int(d))
    except:
        messagebox.showwarning("Заполните дату!","Неправильно заполнена дата!")
    
    input_file = input_file_l['text']
    output_file = filedialog.askdirectory() + input_file_l['text'].split('.')[0] + ".pdf"

    print(output_file)
    s_den = s_den_e.get()
    s_mesyac = s_mesyac_e.get()
    s_god = s_god_e.get()
    po_den = po_den_e.get()
    po_mesyac = po_mesyac_e.get()
    po_god = po_god_e.get()

    processKinds(smena_no,smena_date,input_file,output_file,[s_den,s_mesyac,s_god], [po_den,po_mesyac,po_god])

root = Tk()
root.title("Путёвки")
root.geometry("320x240")
root.resizable(0,0)
#номер смены
Label(root, text="Номер смены:").place(x=10,y=10)
smena_no_e = Entry(root, width=4)
smena_no_e.place(x=110,y=10)

#дата начала смены
Label(root, text="Дата начала смены:").place(x=10,y=35)
smena_date_e = Entry(root, width=10)
smena_date_e.place(x=155,y=35)

#срок путёвки
Label(root,text="Срок путёвки с").place(x=10,y=60)
s_den_e = Entry(root, width=5)
s_den_e.place(x=125,y=60)
s_mesyac_e = Entry(root,width=10)
s_mesyac_e.place(x=175,y=60)
Label(root, text="201").place(x=261,y=60)
s_god_e = Entry(root, width=1)
s_god_e.place(x=285,y=60)
Label(root,text="по").place(x=100,y=80)
po_den_e = Entry(root, width=5)
po_den_e.place(x=125,y=80)
po_mesyac_e = Entry(root,width=10)
po_mesyac_e.place(x=175,y=80)
Label(root, text="201").place(x=261,y=80)
po_god_e = Entry(root, width=1)
po_god_e.place(x=285,y=80)

Button(root,text="Выбрать файл", command=selectFile).place(x=10,y=120)
input_file_l = Label(root,text="Файл не выбран!")
input_file_l.place(x=150,y=120)
Button(root,text="Создать!", command=generate).place(x=110,y=180)

root.mainloop()