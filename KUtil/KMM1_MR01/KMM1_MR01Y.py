import os
import sys
import tkinter as tk
from tkinter import filedialog
import win32com.client as win32
from datetime import datetime, timedelta
import threading
import time


:
import KMM1_MR01B as func_b
import KMM1_MR01C as func_c


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
        self.window.title("자재 요청서 생성")
        #self.window.geometry("670x300+100+100")
        self.window.resizable(False, False)

        self.menubar = MenuBar(self.window, self)
        self.window.config(menu=self.menubar)

        self.dir_base = os.getcwd().replace("/", "\\")
        self.dir_bin  = self.dir_base + "\\_bin"
        self.dir_tmpl = self.dir_base + "\\_tmpl"
        self.dir_curr = os.getcwd().replace("/", "\\")
        self.pdf_name = ''
        self.conv1_name = ''
        self.xmast_name = ''
        self.xdoc_b_name = ''
        self.xdoc_c_name = ''
        self.tmpl_b_name = '_tmpl_반출요청서.xlsx'
        self.tmpl_c_name = '_tmpl_자재요청서.xlsx'

        self.CheckVar21 = False
        self.CheckVar31 = True
        self.CheckVar41 = True
        self.CheckVar51 = True

        self.frame10 = tk.Frame(self.window, width=750, height=60,  padx=4, pady=6, relief='raised', bd=2, bg='black')
        self.frame20 = tk.Frame(self.window, width=750, height=100, padx=4, pady=4, relief='groove', bd=2)
        self.frame30 = tk.Frame(self.window, width=750, height=100, padx=4, pady=4, relief='groove', bd=2)
        self.frame40 = tk.Frame(self.window, width=750, height=100, padx=4, pady=4, relief='groove', bd=2)
        self.frame90 = tk.Frame(self.window, width=750, height=200, padx=4, pady=4, relief='groove', bd=2)

        self.frame11 = tk.Frame(self.frame10, padx=2, pady=0,  bd=0, bg='black')
        self.frame12 = tk.Frame(self.frame10, padx=2, pady=0,  bd=0, bg='black')

        self.frame21 = tk.Frame(self.frame20, padx=2, pady=2, bd=2)
        self.frame22 = tk.Frame(self.frame20, padx=2, pady=2, bd=2)
        self.frame23 = tk.Frame(self.frame20, padx=2, pady=2, bd=2)
        self.frame24 = tk.Frame(self.frame20, padx=2, pady=2, bd=2)
        self.frame25 = tk.Frame(self.frame20, padx=2, pady=2, bd=2)

        self.frame31 = tk.Frame(self.frame30, padx=2, pady=2, bd=2)
        self.frame32 = tk.Frame(self.frame30, padx=2, pady=2, bd=2)
        self.frame33 = tk.Frame(self.frame30, padx=2, pady=2, bd=2)
        self.frame34 = tk.Frame(self.frame30, padx=2, pady=2, bd=2)
        self.frame35 = tk.Frame(self.frame30, padx=2, pady=2, bd=2)

        self.frame41 = tk.Frame(self.frame40, padx=2, pady=2, bd=2)
        self.frame42 = tk.Frame(self.frame40, padx=2, pady=2, bd=2)
        self.frame43 = tk.Frame(self.frame40, padx=2, pady=2, bd=2)
        self.frame44 = tk.Frame(self.frame40, padx=2, pady=2, bd=2)
        self.frame45 = tk.Frame(self.frame40, padx=2, pady=2, bd=2)

        self.label110  = tk.Label (self.frame11, text='자재 요청서 생성', width=27,           padx=24, pady=3, font=('Impact 17', 17, 'bold'), fg='white', bg='black')
        self.label111  = tk.Label (self.frame11, text='',                 width=4,           padx=4, pady=3,                                  bg='black')
        self.button11  = tk.Button(self.frame12, text="SOS",              width=6, height=1, padx=4, pady=0,
                                                 command=lambda: self.zf_close_wb_click())
        self.button12  = tk.Button(self.frame12, text="isExist",          width=6, height=1, padx=4, pady=0,
                                                 command=lambda: self.zs_check_file_click(self.dir_curr))
        self.button13  = tk.Button(self.frame12, text="종료",          width=6, height=1, padx=4, pady=0,
                                                 command=lambda: self. zs_quit_click())

        self.label210  = tk.Label (self.frame21, text='자재청구서(PDF)',  width=14,            padx=4, pady=2)
        self.entry210  = tk.Entry (self.frame22,                          width=50,                           relief='sunken', bg='white')
        self.label211  = tk.Label (self.frame23, text='    ',             width=2,            padx=4, pady=2, relief='raised')
        self.button210 = tk.Button(self.frame24, text="선택",             width=5,  height=2, padx=4, pady=4,
                                                 command=lambda: self.btn_run_click('0'))
        self.label220  = tk.Label (self.frame21, text='작업폴더',         width=14,           padx=4, pady=2)
        self.entry220  = tk.Entry (self.frame22,                          width=50,                            relief='sunken', bg='white')
        self.label221  = tk.Label (self.frame23, text='     ',            width=2,            padx=4, pady=2, relief='raised')
        self.button220 = tk.Button(self.frame25, text="열기",             width=6,  height=2, padx=4, pady=4,
                                                 command=lambda: self.btn_exe_excel_open(self.pdf_name+'.pdf'))

        self.label310  = tk.Label (self.frame31, text='변환(PDF->XLS)',   width=14,           padx=4, pady=2, justify='right')
        self.entry310  = tk.Entry (self.frame32,                          width=50,                           relief='sunken')
        self.label311  = tk.Label (self.frame33, text='      ',           width=2,            padx=4, pady=2, relief='raised')
        self.button310 = tk.Button(self.frame34, text="변환",             width=5,  height=2, padx=4, pady=4,
                                                 command=lambda: self.btn_run_click('A'))
        self.label320  = tk.Label (self.frame31, text='자재청구서',       width=14,           padx=4, pady=2, justify='right')
        self.entry320  = tk.Entry (self.frame32,                          width=50,                           relief='sunken', bg='white')
        self.label321  = tk.Label (self.frame33, text='      ',           width=2,            padx=4, pady=2, relief='raised')
        self.button320 = tk.Button(self.frame35, text="열기",             width=6,  height=1, padx=4, pady=0,
                                                 command=lambda: self.btn_exe_excel_open(self.conv1_name))
        self.button321  = tk.Button(self.frame35, text="열기",           width=6, height=1, padx=4, pady=0,
                                                 command=lambda: self.btn_exe_excel_open(self.xmast_name))

        self.label410  = tk.Label(self.frame41, text='반출요청서',        width=14,           padx=4, pady=2, justify='right')
        self.entry410  = tk.Entry(self.frame42,                         width=50,                           relief='sunken', bg='white')
        self.label411  = tk.Label(self.frame43, text='      ',          width=2,            padx=4, pady=2, relief='raised')
        self.button410 = tk.Button(self.frame44, text="작성",            width=5, height=1,  padx=4, pady=0,
                            command=lambda: self.btn_run_click('B'))
        self.button411 = tk.Button(self.frame45, text="열기",            width=6, height=1,  padx=4, pady=0,
                            command=lambda: self.btn_exe_excel_open(self.xdoc_b_name))

        self.label420  = tk.Label(self.frame41, text='자재요청서',        width=14,           padx=4, pady=2, justify='right')
        self.entry420  = tk.Entry(self.frame42,                         width=50,                           relief='sunken', bg='white')
        self.label421 = tk.Label(self.frame43, text='      ',           width=2,            padx=4, pady=2, relief='raised')
        self.button420 = tk.Button(self.frame44, text="작성",            width=5, height=1,  padx=4, pady=0,
                            command=lambda: self.btn_run_click('C'))
        self.button421 = tk.Button(self.frame45, text="열기",            width=6, height=1,  padx=4, pady=0,
                            command=lambda: self.btn_exe_excel_open(self.xdoc_c_name))

        self.frame10.pack(expand=True, padx=6, pady=6)
        self.frame20.pack(expand=True, padx=6, pady=6)
        self.frame30.pack(expand=True, padx=6, pady=6)
        self.frame40.pack(expand=True, padx=6, pady=6)
        #self.frame90.pack(expand=True)

        self.frame11.pack(side='left')
        self.frame12.pack(side='left')

        self.frame21.pack(side='left')
        self.frame22.pack(side='left')
        self.frame23.pack(side='left')
        self.frame24.pack(side='left')
        self.frame25.pack(side='left')

        self.frame31.pack(side='left')
        self.frame32.pack(side='left')
        self.frame33.pack(side='left')
        self.frame34.pack(side='left')
        self.frame35.pack(side='left')

        self.frame41.pack(side='left')
        self.frame42.pack(side='left')
        self.frame43.pack(side='left')
        self.frame44.pack(side='left')
        self.frame45.pack(side='left')

        self.label110.pack(side='left')
        self.label111.pack(side='left')
        self.button13.pack(side='right')
        self.button11.pack(side='right')
        self.button12.pack(side='right')

        self.label210.pack(anchor='w')
        self.entry210.pack(ipadx=2, ipady=2)
        self.label211.pack()
        self.button210.pack()

        self.label220.pack(anchor='w')
        self.entry220.pack(ipadx=2, ipady=2)
        self.label221.pack()
        self.button220.pack(side='left')

        self.label310.pack(anchor='w')
        self.entry310.pack(ipadx=2, ipady=2)
        self.label311.pack()
        self.button310.pack()

        self.label320.pack(anchor='w')
        self.entry320.pack(ipadx=2, ipady=2)
        self.label321.pack()
        self.button320.pack()
        self.button321.pack()

        self.label410.pack(anchor='w')
        self.entry410.pack(ipadx=2, ipady=2)
        self.label411.pack()
        self.button410.pack()
        self.button411.pack()

        self.label420.pack()
        self.entry420.pack(ipadx=2, ipady=2)
        self.label421.pack()
        self.button420.pack()
        self.button421.pack()

        self.__main__()

    def btn_run_click(self, selfunc):
        bgcolor = 'yellow'

        if selfunc == '0':
            self.label211.config(bg=bgcolor)
            self.label211.update()
            self.label221.config(bg=bgcolor)
            self.label221.update()
            self.btn_get_file_pdf_click()
            #my_thread = threading.Thread(target=self.btn_get_file_pdf_click)

        elif selfunc == 'A':
            self.label311.config(bg=bgcolor)
            self.label311.update()
            self.label321.config(bg=bgcolor)
            self.label321.update()
            self.btn_run_func_a_click()
            #my_thread = threading.Thread(target=self.btn_run_func_a_click)

        elif selfunc == 'B':
            self.label411.config(bg=bgcolor)
            self.label411.update()
            self.btn_run_func_b_click()
            #my_thread = threading.Thread(target=self.btn_run_func_b_click)

        elif selfunc == 'C':
            self.label421.config(bg=bgcolor)
            self.label421.update()
            self.btn_run_func_c_click()
            #my_thread = threading.Thread(target=self.btn_run_func_c_click)

        '''
        my_thread = threading.Thread(target=func_a.zf_pdf_2_xls, args=(input_pdf, output_xls,))

        my_thread.start()
        sval=''
        while my_thread.is_alive():
            time.sleep(0.01)
        '''
        self.zs_check_file_click(self.dir_curr)

    def btn_get_file_pdf_click(self):

        self.file_pdf = self.zf_load_file_pdf()

        self.dir_curr = os.path.dirname(self.file_pdf).replace("/", "\\")
        self.pdf_name = os.path.splitext(os.path.basename(self.file_pdf))[0]
        self.conv1_name = self.pdf_name + '_cnv1.xls'
        self.xmast_name = self.pdf_name + '_mast.xlsx'
        self.xdoc_b_name = self.pdf_name + '_mast_반출요청서.xlsx'
        self.xdoc_c_name = self.pdf_name + '_mast_자재요청서.xlsx'

        self.entry210.delete(0, tk.END)
        self.entry210.insert(0, self.pdf_name)
        self.entry220.delete(0, tk.END)
        self.entry220.insert(0, self.dir_curr)
        self.entry310.delete(0, tk.END)
        self.entry310.insert(0, self.conv1_name)
        self.entry320.delete(0, tk.END)
        self.entry320.insert(0, self.xmast_name)
        self.entry410.delete(0, tk.END)
        self.entry410.insert(0, self.xdoc_b_name)
        self.entry420.delete(0, tk.END)
        self.entry420.insert(0, self.xdoc_c_name)


    def btn_run_func_a_click(self):
        bgcolor = 'blue'

        dircurr = self.dir_curr

        # 파일을 형식 변경 ( PDF --> xlsx )
        input_pdf   = self.dir_curr + '\\' + self.pdf_name + '.pdf'
        output_xls  = self.dir_curr + '\\' + self.conv1_name
        output_xlsx = self.dir_curr + '\\' + self.xmast_name

        func_a.zs_print_message(2, 'converting PDF -> xls')
        func_a.zf_pdf_2_xls(input_pdf, output_xls)

        # 파일을 MS Excel 형식 변경 ( xls --> xlsx )
        func_a.zs_print_message(2, 'converting xls -> xlsx')
        result = func_a.zf_xls_2_xlsx(output_xls , output_xlsx)

    def btn_run_func_b_click(self):

        dircurr = self.dir_curr

        input_xlsx  = self.dir_curr + '\\' + self.xmast_name
        output_xlsx = self.dir_curr + '\\' + self.xdoc_b_name

        # 파일 생성 (자재요청서)
        func_b.zs_print_message(2, 'creating 반출요청서')
        result = func_b.zf_create_carryout(input_xlsx, output_xlsx)


    def btn_run_func_c_click(self):
        bgcolor = 'blue'
        self.label421.config(bg=bgcolor)

        dircurr = self.dir_curr

        input_xlsx = self.dir_curr + '\\' + self.xmast_name
        output_xlsx = self.dir_curr + '\\' + self.xdoc_c_name

        # 파일 생성 (자재요청서)
        func_c.zs_print_message(2, 'creating 자재요청서')
        result = func_c.zf_create_mr(input_xlsx, output_xlsx)


    def zs_check_file_click(self, dir_curr):
        pdfnam = dir_curr + '\\' + self.pdf_name + '.pdf'
        xcnv1  = dir_curr + '\\' + self.conv1_name
        xmast  = dir_curr + '\\' + self.xmast_name
        xdoc_b = dir_curr + '\\' + self.xdoc_b_name
        xdoc_c = dir_curr + '\\' + self.xdoc_c_name

        bgcolor = 'red'
        if os.path.isfile(pdfnam): bgcolor = 'green'
        self.label211.config(bg=bgcolor)
        self.label211.update
        self.label221.config(bg=bgcolor)
        self.label221.update

        bgcolor = 'red'
        if os.path.isfile(xcnv1): bgcolor = 'green'
        self.label311.config(bg=bgcolor)
        self.label311.update

        bgcolor = 'red'
        if os.path.isfile(xmast): bgcolor = 'green'
        self.label321.config(bg=bgcolor)
        self.label321.update

        bgcolor = 'red'
        if os.path.isfile(xdoc_b): bgcolor = 'green'
        self.label411.config(bg=bgcolor)
        self.label411.update

        bgcolor = 'red'
        if os.path.isfile(xdoc_c): bgcolor = 'green'
        self.label421.config(bg=bgcolor)
        self.label421.update


    def btn_exe_excel_open(self, filename):
        filepath = self.dir_curr + '\\' + filename
        file_name, file_ext = os.path.splitext(filename)

        file_ext = file_ext.replace('.', '')
        if file_ext.upper() in ['XLSX', 'XLS']: os.system(f'start excel "{filepath}"')
        if file_ext.upper() in ['PDF']:         os.system(f'start msedge "{filepath}"')


    def zf_close_wb_click(self):

        list_wb = list()
        list_wb.append(self.xmast_name)
        list_wb.append(self.xdoc_b_name)
        list_wb.append(self.xdoc_c_name)
        list_wb.append(self.tmpl_b_name)
        list_wb.append(self.tmpl_c_name)

        com_app = win32.dynamic.Dispatch('Excel.Application')
        com_wbs = com_app.Workbooks
        list_wb_names = [wb.Name for wb in com_wbs]

        self.zs_print_message(2, f'closing {str(list_wb_names)}')

        lcnt = com_wbs.Count
        for i in reversed(range(lcnt)):
            wb_name = com_wbs[i].Name

            for wbnam in list_wb:
                if wbnam in wb_name:
                    self.zs_print_message(2, f'closed {wb_name}')
                    com_wbs[i].Close(SaveChanges=False)
        com_app.Quit()


    def zs_quit_click(self):
        self.window.destroy()


    def zf_load_file_pdf(self):
        filename = filedialog.askopenfilename(initialdir="./", title="Select file",
                                              filetypes=(("PDF files", "*.pdf"),
                                                         ("all files", "*.*")))
        return filename



    def zs_print_message(self, a_sep, a_mesg):
        now = '[' + datetime.now().strftime('%m/%d/%Y-%H:%M:%S') + ']'
        if a_sep == 0:
            print('==========================================================')

        print(now, sys._getframe(1).f_code.co_name + "()", a_mesg, sep=':')

        if a_sep == 9:
            print('----------------------------------------------------------')



    def __main__(self):

        func_c.zs_print_message(0, f' Welcome !!!')
        func_c.zs_print_message(9, f' ')
        #for i in tqdm(range(100), mininterval=1):
        #    print(i, end='')


if __name__ == '__main__':
    window = tk.Tk()

    dir_curr = os.getcwd()
    windows_tkinter(window)
    #window.iconbitmap(f'{dir_curr}\\_image\\KOIN.ico')
    #window.iconbitmap('KOIN.ico')

    ico = tk.PhotoImage(file=f'{dir_curr}\\_image\\KOIN.png')
    window.iconphoto(False, ico)

    window.mainloop()


    """
    my_thread = threading.Thread(target =zf_pdf_2_xls, args=(input_pdf, output_xls,))
    my_thread.start()

    while my_thread.is_alive():
        print(".", end='')
        time.sleep(0.001)
    print('')
    """
