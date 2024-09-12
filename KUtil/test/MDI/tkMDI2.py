'''
    https://pythonassets.com/posts/create-a-new-window-in-tk-tkinter/
https://django.pythonassets.com/docs/templating-system/urls
'''

import tkinter as tk
from tkinter import ttk

def SecondaryWindow1():
    # Create secondary (or popup) window.
    secondary_window = tk.Toplevel()
    secondary_window.title("Secondary Window1")
    secondary_window.config(width=300, height=200)
    # Create a button to close (destroy) this window.
    button_close = ttk.Button(
        secondary_window,
        text="Close window",
        command=secondary_window.destroy
    )
    button_close.place(x=75, y=75)
    secondary_window.focus()

def SecondaryWindow2():
    # Create secondary (or popup) window.
    secondary_window = tk.Toplevel()
    secondary_window.title("Secondary Window2")
    secondary_window.config(width=300, height=200)
    # Create a button to close (destroy) this window.
    button_close = ttk.Button(
        secondary_window,
        text="Close window",
        command=secondary_window.destroy
    )
    button_close.place(x=100, y=75)
    secondary_window.focus()
    secondary_window.grab_set()  # Modal.
class SecondaryWindow3(tk.Toplevel):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config(width=300, height=200)
        self.title("Secondary Window3")
        self.button_close = ttk.Button(
            self,
            text="Close window",
            command=self.destroy
        )
        self.button_close.place(x=100, y=75)
        self.focus()
        self.grab_set()


class MainWindow(tk.Tk):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config(width=400, height=800)
        self.title("Main Window")
        self.button_open1 = ttk.Button(
            self,
            text="Open secondary window",
            command=self.open_secondary_window1
        )
        self.button_open2 = ttk.Button(
            self,
            text="Open secondary window",
            command=self.open_secondary_window2
        )
        self.button_open3 = ttk.Button(
            self,
            text="Open secondary window",
            command=self.open_secondary_window3
        )

        self.button_open1.place(x=100, y=100)
        self.button_open2.place(x=100, y=300)
        self.button_open3.place(x=100, y=500)

    def open_secondary_window1(self):
        self.secondary_window = SecondaryWindow1()
    def open_secondary_window2(self):
        self.secondary_window = SecondaryWindow2()
    def open_secondary_window3(self):
        self.secondary_window = SecondaryWindow3()

main_window = MainWindow()
main_window.mainloop()