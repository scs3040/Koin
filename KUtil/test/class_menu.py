import tkinter as tk
# import Tkinter as tk  # if using python 2
import sys

class MenuBar(tk.Menu):
    def __init__(self, parent):
        super().__init__(parent)

        fileMenu = tk.Menu(self, tearoff=False)
        self.add_cascade(label="File",underline=0, menu=fileMenu)
        fileMenu.add_command(label="Exit", underline=1, command=self.quit)

    def quit(self):
        sys.exit(0)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        menubar = MenuBar(self)
        self.config(menu=menubar)

if __name__ == "__main__":
    app=App()
    app.mainloop()