from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from threading import Thread
from tkinter import ttk 
import tkinter as tk
from SapAutomation import sap_update
from GraphParser import run_update


def parser_amoun():
    get_data = run_update()
    amount = []             
    for i in get_data:
        amount.append(i)
    return amount


class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        notes = ttk.Panedwindow(self)
        notes.grid(column=0, row=0)
        notes.rowconfigure(0, weight=1)
        self.page = ttk.Frame(notes)
        notes.add(self.page)
        self.bind('<Escape>', self.quit)
        self.attributes("-topmost",True)
        self.attributes('-fullscreen', True)
        self.config(cursor="none")
        self.plotter()
        input_frame = ttk.Frame(self)
        input_frame.grid(column=1, row=0)
        self.new_draw()
        Thread(target=self.update_sap_data, daemon=True).start()


    def plotter(self):
        self.figure = Figure(figsize=(20,11), dpi=100)
        self.figure.suptitle("Замовлення з перевищеним терміном виконання більше 2-ох днів")
        self.plot_canvas = FigureCanvasTkAgg(self.figure, self.page)
        self.axes = self.figure.add_subplot(111)
        self.plot_canvas.get_tk_widget().grid(column=0, row=0, sticky='nsew')

    def new_draw(self):
        self.axes.clear()
        self.axes.bar(["Seat", "BMW", "Audi", "Daimler"], parser_amoun(), color='r')
        self.axes.tick_params(axis='both', which='both', labelsize=32)
        self.axes.set_ylabel(' Кількість комплектів   ', fontsize = 24)
        self.plot_canvas.draw_idle()
        self.page.after(1000*5, self.new_draw)

    def update_sap_data(self):
        sap_update()
        self.after(1000*60*15, self.update_sap_data)    
    
    def quit(self, event):
        self.destroy()
        

App().mainloop()
