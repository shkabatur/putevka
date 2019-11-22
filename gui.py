from tkinter import messagebox, filedialog, Entry, Tk, StringVar, Label, END, Button
import re

def selectFile():
    filename = filedialog.askopenfilename(
        initialdir = "/home/den/src/putevka/", title = "Выберитей файл путёвок!",
        filetypes = (("exel files","*.xlsx"), ("all files","*.*")))
    output_file_l['text'] = re.split(r'\\|/', filename)[-1]

def generate():
    print("kek")

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
output_file_l = Label(root,text="Файл не выбран!")
output_file_l.place(x=150,y=120)
Button(root,text="Создать!", command=generate).place(x=110,y=180)


#messagebox.showinfo("KEK", "KEK!")
root.mainloop()