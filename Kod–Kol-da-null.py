#Библиотеки
import tkinter as tk
from tkinter import messagebox
#Чекалка
hod = "X"

def clc(nomer):
    global hod
    knopki[nomer].config(text=hod)
    if hod == "X":
        hod = "O"
    else:
        hod = "X"
        запуск окошка
root = tk.Tk()
root.title("new game")
root.geometry("500x500")
#кнопки 
knopki = []

knopka0 = tk.Button(width=20, height=10, text="", command=lambda: clc(0))
knopka0.place(x=0, y=0)
knopki.append(knopka0)

knopka1 = tk.Button(width=20, height=10, text="", command=lambda: clc(1))
knopka1.place(x=150, y=0)
knopki.append(knopka1)

knopka2 = tk.Button(width=20, height=10, text="", command=lambda: clc(2))
knopka2.place(x=300, y=0)
knopki.append(knopka2)

knopka3 = tk.Button(width=20, height=10, text="", command=lambda: clc(3))
knopka3.place(x=0, y=160)
knopki.append(knopka3)

knopka4 = tk.Button(width=20, height=10, text="", command=lambda: clc(4))
knopka4.place(x=150, y=160)
knopki.append(knopka4)

knopka5 = tk.Button(width=20, height=10, text="", command=lambda: clc(5))
knopka5.place(x=300, y=160)
knopki.append(knopka5)

knopka6 = tk.Button(width=20, height=10, text="", command=lambda: clc(6))
knopka6.place(x=0, y=320)
knopki.append(knopka6)

knopka7 = tk.Button(width=20, height=10, text="", command=lambda: clc(7))
knopka7.place(x=150, y=320)
knopki.append(knopka7)

knopka8 = tk.Button(width=20, height=10, text="", command=lambda: clc(8))
knopka8.place(x=300, y=320)
knopki.append(knopka8)
#запуск окошка (только должен быть внизу) 
root.mainloop()