try:
    import tkinter as tk
    from tkinter import ttk, filedialog
except ImportError:
    import Tkinter as tk
    import ttk, filedialog

from tkcalendar import Calendar, DateEntry

file_to_open = ""
date = ""

s_god = ""
s_mesyac = ""
s_den = ""

po_god = ""
po_mesyac = ""
po_den = ""

def pickDate():
    def setDate():
        file_to_open = cal.selection_get()
        top = tk.Toplevel(root)
        cal = Calendar(top,
                   font="Arial 14", selectmode='day',
                   cursor="hand1", year=2019, month=11, day=5)
        cal.pack(fill="both", expand=True)
        ttk.Button(top, text="ok", command=print_sel).pack()
