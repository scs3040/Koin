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
        self.font17 = tk.font.Font(family="Arial unicode MS", size=17);
        self.font11 = tk.font.Font(family="Arial unicode MS", size=11);
        self.font10 = tk.font.Font(family="Arial unicode MS", size=10);
        self.fields = '폴드 INP', '폴드 OUT'

        self.ent = self.makeform(root, self.fields)

        root.bind('<Return>', (lambda event, e=self.ents: self.fetch(e)))
        row = tk.Frame(root)
        self.b1 = tk.Button(row, text='변환', command=(lambda e=self.ents: self.start_conversion(e)), width=15, height=2,
                            font=self.font11)
        self.b2 = tk.Button(row, text='종료', command=root.quit, width=15, height=2, font=self.font11)
        self.b2.pack(side=tk.RIGHT, padx=5, pady=2)
        self.b1.pack(side=tk.RIGHT, padx=10, pady=2)
        row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        row = tk.Frame(root)
        self.progress_lab = tk.Label(row, text="대기 중...", anchor='w', width=100, height=1, font=self.font11,
                                     relief="ridge")
        self.progress_lab.pack(side=tk.LEFT, expand=tk.YES, padx=14, pady=1)
        row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        row = tk.Frame(root)
        self.progress_ent = tk.Entry(row, font=self.font11, relief="ridge")
        self.progress_ent.pack(side=tk.LEFT, expand=tk.YES, padx=5, pady=1)
        # row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        row = tk.Frame(root)
        self.progress_bar = ttk.Progressbar(row, maximum=100, orient="horizontal", length=660, mode="determinate")
        self.progress_bar.pack(side=tk.LEFT, expand=tk.YES, padx=5, pady=1)
        row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        root.mainloop()

    def fetch(self, entries):
        print(entries[1][0])
        print(entries[1][1].get())
        for entry in entries:
            field = entry[0]
            text = entry[1].get()
            print('%s: "%s"' % (field, text))

    def makeform(self, root, fields):
        entries = []

        row = tk.Frame(root)
        lab = tk.Label(row, width=100, text="변환할 PDF가 있는 폴드를 선택하세요", anchor='w', font=self.font17)
        row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=10)
        lab.pack(side=tk.LEFT)

        for field in self.fields:
            row = tk.Frame(root)
            lab = tk.Label(row, width=10, text=field, anchor='w', font=self.font10)
            ent = tk.Entry(row, font=self.font10)

            row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
            lab.pack(side=tk.LEFT, padx=5)
            ent.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, ipadx=5, ipady=5)
            entries.append((field, ent))

            if field == '폴드 INP':
                btn = tk.Button(row, text=' 선택 ', command=self.get_dir_input, width=10, height=1, font=self.font11)
            elif field == '폴드 OUT':
                btn = tk.Button(row, text=' 선택 ', command=self.get_dir_output, width=10, height=1, font=self.font11)
            else:
                btn = tk.Button(row, text='     ', command='', width=10, height=1)

            btn.pack(side=tk.RIGHT, fill=tk.X, padx=5)
            row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        return entries

    def get_entry(self, sel):
        return self.ents[sel][1].get()

    def set_entry(self, sel, val):
        self.ents[sel][1].delete(0, tk.END)
        self.ents[sel][1].insert(0, val)

    def get_dir_input(self):
        directory = self.select_directory(" PDF 파일 디렉토리를 선택하세요")
        if directory is not None:
            self.set_entry(0, directory)

    def get_dir_output(self):
        directory = self.select_directory(" JPG 파일 저장 디렉토리를 선택하세요")
        if directory is not None:
            self.set_entry(1, directory)

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

        pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
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

    def start_conversion(self, enties):

        input_directory = self.get_entry(0)
        if input_directory == '':
            self.progress_lab.config(text=f' PDF가 있는 폴드를 선택하세요.')

        output_directory = self.get_entry(1)
        if output_directory == '':
            self.progress_lab.config(text=f' PDF를 저장 할 폴드를 선택하세요.')

        self.progress_lab.config(text=" 진행 중...")
        self.progress_bar["value"] = 0
        self.progress_bar.update_idletasks()

        self.convert_directory_pdfs_to_jpg(input_directory, output_directory)


# 메인 프로그램 실행
if __name__ == "__main__":
    root = Tk()
    app = PDFToJPGConverter(root)