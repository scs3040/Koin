import os
import sys
import tkinter as tk
from tkinter import filedialog
import win32com.client as win32
from datetime import datetime, timedelta

import CMM1_1000A as f1000_a

class MenuBar(tk.Menu):
    def __init__(self, parent, wwin):
        super().__init__(parent)

        self.zwin = wwin

        fileMenu = tk.Menu(self, tearoff=False)
        self.add_cascade(label="File",underline=0, menu=fileMenu)
        fileMenu.add_command(label="Exit", underline=1, command=self.mf_onExit)

        toolMenu = tk.Menu(self, tearoff=False)
        self.add_cascade(label="Tool",underline=0, menu=toolMenu)
        toolMenu.add_command(label="Check_File", underline=1, command=self.mf_onCheckfile)
        toolMenu.add_command(label="SOS", underline=1, command=self.mf_onSos)


    def mf_onSos(self):
        self.zwin.zf_close_wb_click()
        
    def mf_onCheckfile(self):
        self.zwin.zs_check_file_click(self.zwin.dir_curr)

    def mf_onExit(self):
        sys.exit(0)


class windows_tkinter:
    def __init__(self, window):
        self.window = window
        self.window.title("사급재 반입/빈출 현황")
        #self.window.geometry("670x300+100+100")
        self.window.resizable(False, False)

        self.menubar = MenuBar(self.window, self)
        self.window.config(menu=self.menubar)

        self.ms_draw_win()

        self.__main__()

    def __main__(self):
        f1000_a.zs_print_message(0, f' Welcome !!!')
        f1000_a.zs_print_message(9, f' ')

    def ms_draw_win(self):
        #====================================================
        self.frame10 = tk.Frame(self.window, width=1000, height=50,  padx=4, pady=6, relief='raised', bd=2, bg='black')
        self.frame20 = tk.Frame(self.window, width=1000, height=100, padx=4, pady=4, relief='groove', bd=2)
        #====================================================
        self.frame10.pack(expand=True, padx=6, pady=6)
        self.frame20.pack(expand=True, padx=6, pady=6)


if __name__ == '__main__':
    window = tk.Tk()

    dir_curr = os.getcwd()
    windows_tkinter(window)

    #ico = tk.PhotoImage(file=f'{dir_curr}\\_image\\KOIN.png')
    #window.iconphoto(False, ico)

    window.mainloop()