# -*- coding:utf-8 -*-
import os
import fitz  # PyMuPDF
from PIL import Image
from tkinter import Tk, Label, Entry, Button, filedialog, ttk
import tkinter as tk
import tkinter.font

class PDFToJPGConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to JPG Converter")
        self.root.geometry("700x700")

        self.ents = []
        self.ents = self.makeform(root)
        print(self.ents)
        root.mainloop()

    def makeform(self, root):
        entries = []

        font17 = tk.font.Font(family="Arial unicode MS", size=17);
        font11 = tk.font.Font(family="Arial unicode MS", size=11);
        font10 = tk.font.Font(family="Arial unicode MS", size=10);

        # 제목
        row0 = tk.Frame(root)
        lab0 = tk.Label(row0, width=100, text="변환할 PDF가 있는 폴드를 선택하세요", anchor='w', font=font17)
        row0.pack(side=tk.TOP, fill=tk.X, padx=5, pady=10)
        lab0.pack(side=tk.LEFT)

        # 폴드 INP
        row1 = tk.Frame(root)
        lab1 = tk.Label(row1, width=10, text='폴드INP', anchor='w', font=font10)
        self.ent1 = tk.Entry(row1, font=font10)
        btn1 = tk.Button(row1, text=' 선택 ', command=self.get_dir_input, width=10, height=1, font=font11)

        row1.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        lab1.pack(side=tk.LEFT, padx=5)
        self.ent1.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, ipadx=5, ipady=5)
        btn1.pack(side=tk.RIGHT, fill=tk.X, padx=5)

        entries.append(("dir_inp", self.ent1))
        
        # 파일 목록- Listbox
        row2 = tk.Frame(root)
        row2.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        scrollbar = tk.Scrollbar(row2)
        scrollbar.pack(side="right", fill="y")
        
        lab21 = tk.Label(row2, width=10, text='파일목록', anchor='w', font=font10)
        self.lst2 = tk.Listbox(row2, selectmode="extended", yscrollcommand=scrollbar.set)
        lab22 = tk.Label(row2, width=10, text='', anchor='w', font=font10)
        lab21.pack(side=tk.LEFT, padx=5)
        self.lst2.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, ipadx=5, ipady=5)
        lab22.pack(side=tk.LEFT, padx=6)
        
        scrollbar.config(command=self.lst2.yview)


        # 폴드 OUT
        row3= tk.Frame(root)
        lab3 = tk.Label(row3, width=10, text='폴드OUT', anchor='w', font=font10)
        self.ent3 = tk.Entry(row3, font=font10)
        btn3 = tk.Button(row3, text=' 선택 ', command=self.get_dir_output, width=10, height=1, font=font11)

        row3.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        lab3.pack(side=tk.LEFT, padx=5)
        self.ent3.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, ipadx=5, ipady=5)
        btn3.pack(side=tk.RIGHT, fill=tk.X, padx=5)

        entries.append(("dir_out", self.ent3))

        # Command Button
        row4 = tk.Frame(root)
        btn41 = tk.Button(row4, text='변환', command=(lambda e=self.ents: self.start_conversion(e)), width=15, height=2,
                            font=font11)
        btn42 = tk.Button(row4, text='종료', command=root.quit, width=15, height=2, font=font11)
        row4.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        btn42.pack(side=tk.RIGHT, padx=5, pady=2)
        btn41.pack(side=tk.RIGHT, padx=10, pady=2)

        # Progress Bar
        row5 = tk.Frame(root)
        self.progress_lab = tk.Label(row5, text="대기 중...", anchor='w', width=100, height=1, font=font11,
                                     relief="ridge")
        row5.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        self.progress_lab.pack(side=tk.LEFT, expand=tk.YES, padx=14, pady=1)

        row6 = tk.Frame(root)
        self.progress_ent = tk.Entry(row6, font=font11, relief="ridge")
        self.progress_ent.pack(side=tk.LEFT, expand=tk.YES, padx=5, pady=1)
        # row6.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        row7 = tk.Frame(root)
        self.progress_bar = ttk.Progressbar(row7, maximum=100, orient="horizontal", length=660, mode="determinate")
        row7.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        self.progress_bar.pack(side=tk.LEFT, expand=tk.YES, padx=5, pady=1)
        
        return entries

    def do_entry(self, opt, idx, val=""):
        if opt.upper() == 'GET' :
            return self.ents[idx][1].get()
        
        elif opt.upper() == 'SET' :
            self.ents[idx][1].delete(0, tk.END)
            self.ents[idx][1].insert(0, val)
            return 1
        else:
            return 0

    def get_dir_input(self):
        directory = self.select_files(" PDF 파일 디렉토리를 선택하세요")
        if directory is not None:
            self.do_entry('SET', 0, directory)
            self.do_entry('SET', 1, directory)

    def get_dir_output(self):
        directory = self.select_directory(" JPG 파일 저장 디렉토리를 선택하세요")
        if directory is not None:
            self.do_entry('SET', 1, directory)

    def select_files(self, promot):
        Tk().withdraw()  # Tkinter 기본 윈도우 숨기기
        
        self.lst2.delete(0, tk.END)
        file_paths = filedialog.askopenfilenames(title="파일 선택", filetypes=[("PDF파일", "*.PDF")])
        
        if len(file_paths) > 0:
            for path in file_paths:
                self.lst2.insert(tk.END, os.path.basename(path))
            return os.path.dirname(file_paths[0])
        else:
            return None
    
    def select_directory(self, prompt):
        Tk().withdraw()  # Tkinter 기본 윈도우 숨기기
        directory = filedialog.askdirectory(title=prompt)
        if directory:
            return directory
        else:
            return None

    def convert_pdf_to_jpg(self, pdf_path, output_dir, step):
        pdf_document = fitz.open(pdf_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            pix = page.get_pixmap()
            output_path = os.path.join(
                output_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page_{page_num + 1}.jpg"
            )
            pix.save(output_path)
        pdf_document.close()
        self.progress_bar.step(step)
        self.progress_bar.update()

    def convert_directory_pdfs_to_jpg(self, input_dir, output_dir):
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        pdf_files = [f for f in self.lst2.get(0,tk.END) if f.lower().endswith('.pdf')]
        total_files = len(pdf_files)

        if total_files == 0:
            self.progress_lab.config(text=" PDF 파일이 없습니다.")
            return

        step = 100 / total_files
        for idx, filename in enumerate(pdf_files):
            pdf_path = os.path.join(input_dir, filename)
            self.convert_pdf_to_jpg(pdf_path, output_dir, step)
            self.progress_lab.config(text=f" 파일 변환 완료... {idx + 1} / {total_files} ")
            self.progress_lab.update_idletasks()

        self.progress_lab.config(text=" 변환이 완료되었습니다!")

    def start_conversion(self, entries):
        input_directory = self.ent1.get()
        output_directory = self.ent3.get()

        self.progress_lab.config(text=" 진행 중...")
        self.progress_bar["value"] = 0
        self.progress_bar.update_idletasks()

        self.convert_directory_pdfs_to_jpg(input_directory, output_directory)


# 메인 프로그램 실행
if __name__ == "__main__":
    root = Tk()
    app = PDFToJPGConverter(root)