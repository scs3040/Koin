import sys
import tkinter as tk
from tkinter.simpledialog import askstring


class windows_tkinter:
    def __init__(self, window):
        self.window = window
        self.window.title("자재 요청서 생성")
        #self.window.geometry("670x300+100+100")
        self.window.resizable(False, False)

        self.__main__()

    def __main__(self):

        ms_print_message(0, f' Welcome !!!')
        ms_print_message(9, f' ')

    def zs_print_message(self, a_sep, a_mesg):
        now = '[' + datetime.now().strftime('%m/%d/%Y-%H:%M:%S') + ']'
        if a_sep == 0:
            print('==========================================================')

        print(now, sys._getframe(1).f_code.co_name + "()", a_mesg, sep=':')

        if a_sep == 9:
            print('----------------------------------------------------------')

if __name__ == "__main__":
    app=App()
    app.mainloop()