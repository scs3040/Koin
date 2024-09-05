import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
class windows_tkinter:
    def __init__(self, window):
        self.window = window
        self.window.title("자재청구(반출) 요청서 생성")
        self.window.geometry("800x500+100+100")
        self.window.resizable(False, False)

        self.fname = ''

        self.tklabel=tk.Label(window, text="파이썬", width=10, height=5, fg='white', relief="solid")
        self. tklabel.pack()

        self.__main__()
    def __main__(self):
        dir_curr = os.getcwd()
        dir_parent = os.path.dirname(os.getcwd())

        fname = dir_parent + '\\KMM1_MR01\\' + 'sample.pdf'
        self.fname = fname

        print(fname)
        if self.zf_isfile(fname):
            print(f'a {fname}')
            self.zs_style_label(lable=self.tklabel, bg='green')
        else:
            print(f'b {fname}')
            self.zs_style_label(lable=self.tklabel, bg='red')

    def zf_isfile(self, fname):
        return os.path.isfile(fname)

    def zs_style_label(self, lable, bg, **kwargs):
        lable.config(bg = bg)


if __name__ == '__main__':
    window = tk.Tk()
    windows_tkinter(window)
    window.mainloop()
