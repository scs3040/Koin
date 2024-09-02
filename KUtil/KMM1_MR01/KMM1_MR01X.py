import tkinter
class windows_tkinter:
    def __init__(self, window):
        self.window = window
        self.window.title("자재 청구/불출요청서 생성")
        self.window.geometry("600x300+700+100")
        self.window.resizable(False, False)

        self.arg1 = 1
        self.arg2 = "alpha"
        self.arg3 = "beta"
        self.__main__()

    def command_args(self, argument1, argument2, argument3):
        print(argument1, argument2, argument3)
        self.arg1 = argument1 * 2
        self.label1.text = self.arg1
    def __main__(self):
        frame10 = tkinter.Frame(self.window,  width=580, height=200, relief='groove', bd=2)

        frame11 = tkinter.Frame(frame10, width=140, height=120, padx=2, pady=2, relief='groove', bd=2)
        frame12 = tkinter.Frame(frame10, width=400, height=120, padx=2, pady=2, relief='groove', bd=2)

        label11 = tkinter.Label(frame11, text='자재청구서(PDF)', padx=2, pady=2, relief='groove')
        label12 = tkinter.Label(frame11, text='자재청구서(PDF)', padx=2, pady=2, relief='groove')
        entry11 = tkinter.Label(frame12, width=30, padx=2, pady=2, relief='sunken', bg='white')
        entry12 = tkinter.Label(frame12, width=30, padx=2, pady=2, relief='sunken', bg='white')

        frame10.pack(expand=True, side='top')

        frame11.place(x=0, y=0)
        frame12.place(x=120, y=0)

        label11 = label11.pack()
        label12 = label12.pack()
        entry11 = entry11.pack()
        entry12 = entry12.pack()

        '''
        button11 = tkinter.Button(frame2, text="버튼", width=5, height=1, command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        button11.pack(side='left', padx=1)
        button11 = tkinter.Button(frame2, text="버튼", width=5, height=1, command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        button11.pack(side='left', padx=1)
        '''
if __name__ == '__main__':
    window = tkinter.Tk()
    windows_tkinter(window)
    window.mainloop()
