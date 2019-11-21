try:
    import tkinter as tk
    from tkinter import ttk, filedialog
except ImportError:
    import Tkinter as tk
    import ttk, filedialog

from tkcalendar import Calendar, DateEntry

file_to_open = ""
start_smena = ""

s_god = ""
s_mesyac = ""
s_den = ""

po_god = ""
po_mesyac = ""
po_den = ""


def example1():
    def print_sel():
        global start_smena
        start_smena = cal.selection_get()

    top = tk.Toplevel(root)

    cal = Calendar(top,
                   font="Arial 14", selectmode='day',
                   cursor="hand1", year=2019, month=11, day=5)
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()


def openFileDialog():
    global file_to_open
    file_to_open = filedialog.askopenfilename(initialdir = "/home",title = "Select file",filetypes = (("exel files","*.xlsx"),("all files","*.*")))
    open_file_string.set(file_to_open)
    print(file_to_open)

root = tk.Tk()
s = ttk.Style(root)
s.theme_use('clam')
open_file_string = tk.StringVar()
tk.Label(root, text=open_file_string).place(x = 0, y = 140)
tk.Entry(root, text="KEK:").place(x=0,y=0)
ttk.Button(root, text="OpenFile", command = openFileDialog).place(x=0,y=20)
ttk.Button(root, text='Calendar', command=example1).place(x=0,y=60)

root.mainloop()