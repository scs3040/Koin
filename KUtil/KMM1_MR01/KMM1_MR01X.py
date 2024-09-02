import tkinter
class windows_tkinter:
    def __init__(self, window):
        self.window = window
        self.window.title("자재 청구/불출요청서 생성")
        self.window.geometry("850x800+100+100")
        self.window.resizable(False, False)

        self.arg1 = 1
        self.arg2 = "alpha"
        self.arg3 = "beta"
        self.__main__()

    def command_args(self, argument1, argument2, argument3):
        print(argument1, argument2, argument3)
        self.arg1 = argument1 * 2
        self.label1.text = self.arg1


    def zs_draw_screen(self):
        frame10 = tkinter.Frame(self.window, width=840, height=60, relief='groove', bd=2)
        frame20 = tkinter.Frame(self.window, width=840, height=98, padx=5, pady=5, relief='groove', bd=2)
        frame30 = tkinter.Frame(self.window, width=840, height=100, padx=5, pady=5, relief='groove', bd=2)
        frame40 = tkinter.Frame(self.window, width=840, height=100, padx=5, pady=5, relief='groove', bd=2)
        frame50 = tkinter.Frame(self.window, width=840, height=100, padx=5, pady=5, relief='groove', bd=2)
        frame60 = tkinter.Frame(self.window, width=840, height=300, relief='groove', bd=2)

        frame21 = tkinter.Frame(frame20, width=140, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame22 = tkinter.Frame(frame20, width=600, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame23 = tkinter.Frame(frame20, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame24 = tkinter.Frame(frame20, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame25 = tkinter.Frame(frame20, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame31 = tkinter.Frame(frame30, width=140, height=50, padx=5, pady=6, relief='groove', bd=2)
        frame32 = tkinter.Frame(frame30, width=600, height=50, padx=5, pady=6, relief='groove', bd=2)
        frame33 = tkinter.Frame(frame30, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame34 = tkinter.Frame(frame30, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame35 = tkinter.Frame(frame30, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame41 = tkinter.Frame(frame40, width=140, height=90, padx=5, pady=6, relief='groove', bd=2)
        frame42 = tkinter.Frame(frame40, width=600, height=50, padx=5, pady=6, relief='groove', bd=2)
        frame43 = tkinter.Frame(frame40, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame44 = tkinter.Frame(frame40, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame45 = tkinter.Frame(frame40, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame51 = tkinter.Frame(frame50, width=140, height=50, padx=5, pady=6, relief='groove', bd=2)
        frame52 = tkinter.Frame(frame50, width=600, height=50, padx=5, pady=6, relief='groove', bd=2)
        frame53 = tkinter.Frame(frame50, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame54 = tkinter.Frame(frame50, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)
        frame55 = tkinter.Frame(frame50, width=50, height=50, padx=5, pady=5, relief='groove', bd=2)

        label21 = tkinter.Label(frame21, text='자재청구서(PDF)', width=13, padx=5, pady=6, relief='groove')
        label22 = tkinter.Label(frame21, text='홈디렉토리', width=13, padx=5, pady=6, relief='groove')
        entry21 = tkinter.Label(frame22, width=50, padx=5, pady=6, relief='sunken', bg='white')
        entry22 = tkinter.Label(frame22, width=50, padx=5, pady=6, relief='sunken', bg='white')
        button21 = tkinter.Button(frame23, text="버튼", width=5, height=1, padx=5, pady=2,
                                  command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        button12 = tkinter.Button(frame23, text="버튼", width=5, height=1, padx=5, pady=2,
                                  command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        CheckVar21 = 0
        CheckVar22 = 1
        checkbutton21 = tkinter.Checkbutton(frame24, text='자동', height=1, padx=5, pady=3, variable=CheckVar21)
        checkbutton22 = tkinter.Checkbutton(frame24, text='자동', height=1, padx=5, pady=3, variable=CheckVar22)

        label31 = tkinter.Label(frame31, text='자재청구서', width=13, padx=5, pady=6, relief='groove')
        label32 = tkinter.Label(frame31, text='Master(XLSX)', width=13, padx=5, pady=6, relief='groove')
        entry31 = tkinter.Label(frame32, width=50, padx=5, pady=6, relief='sunken', bg='white')
        entry32 = tkinter.Label(frame32, width=50, padx=5, pady=6, relief='sunken', bg='white')
        button31 = tkinter.Button(frame33, text="버튼", width=5, height=1, padx=5, pady=2,
                                  command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        button32 = tkinter.Button(frame33, text="버튼", width=5, height=1, padx=5, pady=2,
                                  command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        CheckVar31 = 0
        CheckVar32 = 1
        checkbutton31 = tkinter.Checkbutton(frame34, text='자동', height=1, padx=5, pady=3, variable=CheckVar31)
        checkbutton32 = tkinter.Checkbutton(frame34, text='자동', height=1, padx=5, pady=3, variable=CheckVar32)

        label41 = tkinter.Label(frame41, text='템플릿', width=13, padx=5, pady=6, relief='groove')
        label42 = tkinter.Label(frame41, text='자재요청서', width=13, padx=5, pady=6, relief='groove')
        entry41 = tkinter.Label(frame42, width=50, padx=5, pady=6, relief='sunken', bg='white')
        entry42 = tkinter.Label(frame42, width=50, padx=5, pady=6, relief='sunken', bg='white')
        button41 = tkinter.Button(frame43, text="작성", width=5, height=1, padx=5, pady=2,
                                  command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        button42 = tkinter.Button(frame43, text="버튼", width=5, height=1, padx=5, pady=2,
                                  command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        CheckVar41 = 0
        CheckVar42 = 1
        checkbutton41 = tkinter.Checkbutton(frame44, text='자동', height=1, padx=5, pady=3, variable=CheckVar41)
        checkbutton42 = tkinter.Checkbutton(frame44, text='자동', height=1, padx=5, pady=3, variable=CheckVar42)

        label51 = tkinter.Label(frame51, text='템플릿', width=13, padx=5, pady=6, relief='groove')
        label52 = tkinter.Label(frame51, text='불출요청서', width=13, padx=5, pady=6, relief='groove')
        entry51 = tkinter.Label(frame52, width=50, padx=5, pady=6, relief='sunken', bg='white')
        entry52 = tkinter.Label(frame52, width=50, padx=5, pady=6, relief='sunken', bg='white')
        button51 = tkinter.Button(frame53, text="작성", width=5, height=1, padx=5, pady=2,
                                  command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        button52 = tkinter.Button(frame53, text="버튼", width=5, height=1, padx=5, pady=2,
                                  command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        CheckVar51 = 0
        CheckVar52 = 1
        checkbutton51 = tkinter.Checkbutton(frame54, text='자동', height=1, padx=5, pady=3, variable=CheckVar51)
        checkbutton52 = tkinter.Checkbutton(frame54, text='자동', height=1, padx=5, pady=3, variable=CheckVar52)

        frame10.pack(expand=True)
        frame20.pack(expand=True)
        frame30.pack(expand=True)
        frame40.pack(expand=True)
        frame50.pack(expand=True)
        frame60.pack(expand=True)

        frame21.place(x=0, y=0)
        frame22.place(x=140, y=0)
        frame23.place(x=620, y=0)
        frame24.place(x=700, y=0)
        frame25.place(x=780, y=0)
        frame31.place(x=0, y=0)
        frame32.place(x=140, y=0)
        frame33.place(x=620, y=0)
        frame34.place(x=700, y=0)
        frame35.place(x=780, y=0)
        frame41.place(x=0, y=0)
        frame42.place(x=140, y=0)
        frame43.place(x=620, y=0)
        frame44.place(x=700, y=0)
        frame45.place(x=780, y=0)
        frame51.place(x=0, y=0)
        frame52.place(x=140, y=0)
        frame53.place(x=620, y=0)
        frame54.place(x=700, y=0)
        frame55.place(x=780, y=0)

        label21 = label21.pack()
        label22 = label22.pack()
        entry21 = entry21.pack()
        entry22 = entry22.pack()
        button21 = button21.pack()
        # button22 = button22.pack()
        # checkbutton21 = checkbutton21.pack()
        # checkbutton22 = checkbutton22.pack()
        label31 = label31.pack()
        label32 = label32.pack()
        entry31 = entry31.pack()
        entry32 = entry32.pack()
        button31 = button31.pack()
        # button32 = button32.pack()
        checkbutton31 = checkbutton31.pack()
        # checkbutton32 = checkbutton32.pack()
        label41 = label41.pack()
        label42 = label42.pack()
        entry41 = entry41.pack()
        entry42 = entry42.pack()
        button41 = button41.pack()
        # button42 = button42.pack()
        checkbutton41 = checkbutton41.pack()
        # checkbutton42 = checkbutton42.pack()
        label51 = label51.pack()
        label52 = label52.pack()
        entry51 = entry51.pack()
        entry52 = entry52.pack()
        button51 = button51.pack()
        # button52 = button52.pack()
        checkbutton51 = checkbutton51.pack()
        # checkbutton52 = checkbutton52.pack()
        '''
        button11 = tkinter.Button(frame2, text="버튼", width=5, height=1, command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        button11.pack(side='left', padx=1)
        button11 = tkinter.Button(frame2, text="버튼", width=5, height=1, command=lambda: self.command_args(self.arg1, self.arg2, self.arg3))
        button11.pack(side='left', padx=1)
        '''
    def __main__(self):
        self.zs_draw_screen()


if __name__ == '__main__':
    window = tkinter.Tk()
    windows_tkinter(window)
    window.mainloop()
