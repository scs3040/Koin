import os
import sys
import tkinter as tk
from tkinter import filedialog
import win32com.client as win32
from datetime import datetime, timedelta
import threading
import time


import KMM1_MR01A as func_a
import KMM1_MR01B as func_b
import KMM1_MR01C as func_c

class windows_tkinter:
    def __init__(self, window):
        self.window = window
        self.window.title("자재청구(반출) 요청서 생성")
        self.window.geometry("800x500+100+100")
        self.window.resizable(False, False)

        self.dir_home = os.getcwd().replace("/", "\\")
        self.dir_bin  = self.dir_home + "\\_bin"
        self.dir_tmpl = self.dir_home + "\\_tmpl"
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

        self.frame10 = tk.Frame(self.window, width=750, height=60, padx=4, pady=4, relief='groove', bd=2)
        self.frame20 = tk.Frame(self.window, width=750, height=100, padx=4, pady=4, relief='groove', bd=2)
        self.frame30 = tk.Frame(self.window, width=750, height=100, padx=4, pady=4, relief='groove', bd=2)
        self.frame40 = tk.Frame(self.window, width=750, height=100, padx=4, pady=4, relief='groove', bd=2)
        self.frame90 = tk.Frame(self.window, width=750, height=200, padx=4, pady=4, relief='groove', bd=2)

        self.frame11 = tk.Frame(self.frame10, padx=2, pady=2, relief='groove', bd=2)

        self.frame21 = tk.Frame(self.frame20, padx=2, pady=2, relief='groove', bd=2)
        self.frame22 = tk.Frame(self.frame20, padx=2, pady=2, relief='groove', bd=2)
        self.frame23 = tk.Frame(self.frame20, padx=2, pady=2, relief='groove', bd=2)
        self.frame24 = tk.Frame(self.frame20, padx=2, pady=2, relief='groove', bd=2)

        self.frame31 = tk.Frame(self.frame30, padx=2, pady=2, relief='groove', bd=2)
        self.frame32 = tk.Frame(self.frame30, padx=2, pady=2, relief='groove', bd=2)
        self.frame33 = tk.Frame(self.frame30, padx=2, pady=2, relief='groove', bd=2)
        self.frame34 = tk.Frame(self.frame30, padx=2, pady=2, relief='groove', bd=2)

        self.frame41 = tk.Frame(self.frame40, padx=2, pady=2, relief='groove', bd=2)
        self.frame42 = tk.Frame(self.frame40, padx=2, pady=2, relief='groove', bd=2)
        self.frame43 = tk.Frame(self.frame40, padx=2, pady=2, relief='groove', bd=2)
        self.frame44 = tk.Frame(self.frame40, padx=2, pady=2, relief='groove', bd=2)

        self.label11 = tk.Label(self.frame11, text='자재청구(불출) 요청서 생성', padx=1, pady=1, relief='groove', font=("Arial", 15))
        self.button11 = tk.Button(self.frame11, text="SOS", width=8, height=1, padx=1, pady=1,
                            command=lambda: self.zf_close_wb_click())
        self.label111 = tk.Label(self.frame11, text='', width=4, padx=4, pady=5, relief='groove')

        self.label21  = tk.Label(self.frame21, text='자재청구서(PDF)', width=14,           padx=4, pady=2, relief='groove')
        self.entry21  = tk.Entry(self.frame22,                       width=50,                           relief='sunken', bg='white')
        self.label211 = tk.Label(self.frame23, text='    ',          width=2,            padx=4, pady=2, relief='groove')
        self.button21 = tk.Button(self.frame24, text="PDF선택",       width=14, height=2, padx=4, pady=0,
                            command=lambda: self.btn_get_file_pdf_click())
        self.label22  = tk.Label(self.frame21, text='작업폴더',        width=14,           padx=4, pady=2, relief='groove')
        self.entry22  = tk.Entry(self.frame22,                       width=50,                            relief='sunken', bg='white')
        self.label221 = tk.Label(self.frame23, text='     ',          width=2,            padx=4, pady=2, relief='groove')
        self.button22 = tk.Button(self.frame24, text="---",           width=14, height=2, padx=4, pady=0,
                            command=lambda: self.btn_get_file_pdf_click())

        self.label31  = tk.Label (self.frame31, text='변환(PDF->XLSX)', width=14,           padx=4, pady=2, relief='groove')
        self.entry31  = tk.Entry (self.frame32,                        width=50,                           relief='sunken', bg='white')
        self.label311 = tk.Label (self.frame33, text='      ',         width=2,            padx=4, pady=2, relief='groove')
        self.button31 = tk.Button(self.frame34, text="변환(PDF->XLSX)", width=14, height=2, padx=4, pady=0,
                            command=lambda: self.btn_run_func_a_click())
        self.label32  = tk.Label (self.frame31, text='변환(XLS->XLSX)', width=14,           padx=4, pady=2, relief='groove')
        self.entry32  = tk.Entry (self.frame32,                        width=50,                           relief='sunken', bg='white')
        self.label321 = tk.Label (self.frame33, text='      ',         width=2,            padx=4, pady=2, relief='groove')
        self.button32 = tk.Button(self.frame34, text="---",            width=14, height=2, padx=4, pady=0,
                            command=lambda: self.btn_get_file_pdf_click())

        self.label41  = tk.Label(self.frame41, text='반출요청서',        width=14,           padx=4, pady=2, relief='groove')
        self.entry41  = tk.Entry(self.frame42,                         width=50,                          relief='sunken', bg='white')
        self.label411 = tk.Label(self.frame43, text='      ',          width=2,            padx=4, pady=2, relief='groove')
        self.button41 = tk.Button(self.frame44, text="작  성",          width=14, height=1, padx=4, pady=0,
                            command=lambda: self.btn_run_func_b_click())
        self.label42  = tk.Label(self.frame41, text='자재요청서',        width=14,           padx=4, pady=2, relief='groove')
        self.entry42  = tk.Entry(self.frame42,                         width=50,                          relief='sunken', bg='white')
        self.label421 = tk.Label(self.frame43, text='      ',          width=2,            padx=4, pady=2, relief='groove')
        self.button42 = tk.Button(self.frame44, text="작  성",          width=14, height=1, padx=4, pady=0,
                            command=lambda: self.btn_run_func_c_click())

        self.frame10.pack(expand=True)
        self.frame20.pack(expand=True)
        self.frame30.pack(expand=True)
        self.frame40.pack(expand=True)
        self.frame90.pack(expand=True)

        self.frame11.pack(expand=True)

        self.frame21.pack(side='left')
        self.frame22.pack(side='left')
        self.frame23.pack(side='left')
        self.frame24.pack(side='left')

        self.frame31.pack(side='left')
        self.frame32.pack(side='left')
        self.frame33.pack(side='left')
        self.frame34.pack(side='left')

        self.frame41.pack(side='left')
        self.frame42.pack(side='left')
        self.frame43.pack(side='left')
        self.frame44.pack(side='left')

        self.label11.pack(side='left')
        self.button11.pack(side='right')
        self.label111.pack(side='left')
        #self.button12.pack(side='right')

        self.label21.pack()
        self.entry21.pack(ipadx=2, ipady=2)
        self.label211.pack()
        self.button21.pack()

        self.label22.pack()
        self.entry22.pack(ipadx=2, ipady=2)
        self.label221.pack()
        #self.button22.pack(side='left')

        self.label31.pack()
        self.entry31.pack(ipadx=2, ipady=2)
        self.label311.pack()
        self.button31.pack()

        self.label32.pack()
        self.entry32.pack(ipadx=2, ipady=2)
        self.label321.pack()

        self.label41.pack()
        self.entry41.pack(ipadx=2, ipady=2)
        self.label411.pack()
        self.button41.pack()

        self.label42.pack()
        self.entry42.pack(ipadx=2, ipady=2)
        self.label421.pack()
        self.button42.pack()

        self.__main__()


    def btn_get_file_pdf_click(self):
        self.file_pdf = self.zf_load_file_pdf()

        self.dir_curr = os.path.dirname(self.file_pdf).replace("/", "\\")
        self.pdf_name = os.path.splitext(os.path.basename(self.file_pdf))[0]
        self.conv1_name = self.pdf_name + '_cnv1.xls'
        self.xmast_name = self.pdf_name + '_mast.xlsx'
        self.xdoc_b_name = self.pdf_name + '_mast_반출요청서.xlsx'
        self.xdoc_c_name = self.pdf_name + '_mast_자재요청서.xlsx'


        self.entry21.delete(0, tk.END)
        self.entry21.insert(0, self.pdf_name)
        self.entry22.delete(0, tk.END)
        self.entry22.insert(0, self.dir_curr)
        self.entry31.delete(0, tk.END)
        self.entry31.insert(0, self.conv1_name)
        self.entry32.delete(0, tk.END)
        self.entry32.insert(0, self.xmast_name)
        self.entry41.delete(0, tk.END)
        self.entry41.insert(0, self.xdoc_b_name)
        self.entry42.delete(0, tk.END)
        self.entry42.insert(0, self.xdoc_c_name)

    def btn_run_click(self):
        self.zf_run_func1()


    def btn_run_func_a_click(self):
        dircurr = self.dir_curr

        # 파일을 형식 변경 ( PDF --> xlsx )
        input_pdf  = self.dir_curr + '\\' + self.pdf_name + '.pdf'
        output_xls = self.dir_curr + '\\' + self.conv1_name
        output_xlsx = self.dir_curr + '\\' + self.xmast_name

        func_a.zs_print_message(2, 'converting PDF -> xls')

        #func_a.zf_pdf_2_xls(input_pdf, output_xls)
        my_thread = threading.Thread(target=func_a.zf_pdf_2_xls, args=(input_pdf, output_xls,))
        my_thread.start()
        sval=''
        while my_thread.is_alive():
            sval = sval + '**'
            self.label321.config(text=sval)
            #print("*", end='')
            time.sleep(0.0001)
        #print('')

        # 파일을 MS Excel 형식 변경 ( xls --> xlsx )
        func_a.zs_print_message(2, 'converting xls -> xlsx')
        result = func_a.zf_xls_2_xlsx(output_xls , output_xlsx)

        #subprocess(["python", "KMM1_MR01A.py"])


    def btn_run_func_b_click(self):
        dircurr = self.dir_curr

        input_xlsx = self.dir_curr + '\\' + self.xmast_name
        output_xlsx = self.dir_curr + '\\' + self.xdoc_b_name

        # 파일 생성 (자재요청서)
        func_b.zs_print_message(2, 'creating 반출요청서')
        result = func_b.zf_create_carryout(input_xlsx, output_xlsx)


    def btn_run_func_c_click(self):
        dircurr = self.dir_curr

        input_xlsx = self.dir_curr + '\\' + self.xmast_name
        output_xlsx = self.dir_curr + '\\' + self.xdoc_c_name

        # 파일 생성 (자재요청서)
        func_c.zs_print_message(2, 'creating 자재요청서')
        result = func_c.zf_create_mr(input_xlsx, output_xlsx)

    def zf_close_wb_click(self):

        list_wb = list()
        list_wb.append(self.xmast_name)
        list_wb.append(self.xdoc_b_name)
        list_wb.append(self.xdoc_c_name)
        list_wb.append(self.tmpl_b_name)
        list_wb.append(self.tmpl_c_name)
        list_wb.append('Tmpl_반출요청서.xlsx')

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
        print('qq')
        #for i in tqdm(range(100), mininterval=1):
        #    print(i, end='')


if __name__ == '__main__':
    window = tk.Tk()
    windows_tkinter(window)
    window.mainloop()


    """
    my_thread = threading.Thread(target =zf_pdf_2_xls, args=(input_pdf, output_xls,))
    my_thread.start()

    while my_thread.is_alive():
        print(".", end='')
        time.sleep(0.001)
    print('')
    """
