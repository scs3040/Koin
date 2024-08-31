import os
import sys
from datetime import datetime
#from tkinter import *
from tkinter import filedialog

import aspose.pdf as ap
#import asposecells

import win32com.client as win32
from tqdm import tqdm
import xlwings as xw

#import threading
#import time

def zf_FileLoad():
    print_message(2, 'select file ...')
    filename = filedialog.askopenfilename(initialdir="./", title="Select file",
                                          filetypes=(("PDF files", "*.pdf"),
                                                     ("all files", "*.*")))
    print_message(2, 'selected ' + filename)
    return filename


def zf_FileSave():
    print_message(2, 'select file ... ')
    filename = filedialog.asksaveasfilename(initialdir="./", title="Select file",
                                            filetypes=(("XLS files", "*.xls"),
                                                       ("all files", "*.*")))
    # print(filename)
    print_message(2, 'Saved ' + filename)
    return filename


def zf_Xls2Xlsx(input_xls):
    print_message(0, 'Starting ...')

    FileDir = os.path.dirname(input_xls).replace("/", "\\")
    FileNam, FileExt = os.path.splitext(os.path.basename(input_xls))
    output_xlsx = FileDir + "\\" + FileNam + "_1.xlsx"

    print_message(2, 'open PDF ' + input_xls)
    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        wb = excel_app.Workbooks.Open(input_xls)
    except:
        wb.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)

    zs_SheetMerge(wb)

    print_message(2, 'saving ... ' + output_xlsx)
    try:
        wb.SaveAs(output_xlsx, FileFormat=51)
    except:
        print_message(9, 'save cancel ')
        wb.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)

    print_message(2, 'saved      ' + output_xlsx)
    wb.Close()

    excel_app.Application.Quit()

    print_message(9, 'Finished')

    return output_xlsx


def zs_SheetMerge(aWorkbook):
    wb = aWorkbook
    print_message(2, 'starting ...')

    wsCnt = wb.Sheets.Count
    wst = wb.Sheets(1)
    for i in tqdm(range(2, wsCnt + 1), mininterval=1):

        wss = wb.Sheets(i)

        lastrow = wst.UsedRange.Rows.Count

        for i in range(lastrow, 0, -1):
            str1 = wss.Cells(i, 1).Value
            if str1 is None:
                wss.Rows(i).EntireRow.Delete()
            else:
                str2 = str1[0:2]
                if not str2.isnumeric():
                    wss.Rows(i).EntireRow.Delete()

        wss.UsedRange.Copy()

        lastrow = wst.UsedRange.Rows.Count - 2
        wst.Cells(lastrow, 1).Select()
        wst.Paste()
        wst.Cells(1, 1).Value = ''

    print_message(2, 'finished')
    zs_SheetStyle(wst, "A:Z")
def zs_SheetStyle(aWorksheet, aRange ):
    print_message(2, 'starting .. ')

    ws = aWorksheet
    rng = aRange

    ws.Range(aRange).Font.Size = 10

    print_message(2, 'finished')


def zf_MR_create(input_xlsx):
    print_message(0, 'starting ...')

    FileDir = os.path.dirname(input_xlsx).replace("/", "\\")
    FileNam, FileExt = os.path.splitext(os.path.basename(input_xlsx))

    CurrDir = os.getcwd()
    FileTmpl = CurrDir + "\\_Tmpl\\tmpl_자재요청.xlsx"

    output_xlsx = FileDir + "\\" + FileNam + "_자재요청서.xlsx"

    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        wbs = excel_app.Workbooks.Open(input_xlsx)
    except:
        wbs.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)
    try:
        wbt = excel_app.Workbooks.Open(FileTmpl)
    except:
        wbt.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)

    wss = wbs.Sheets(1)
    srows = wss.Range("A:A").Find('품번', LookAt=1).Row + 1
    srowf = get_last_row_from_column(wss, 'A')

    wst = wbt.Sheets(1)
    trows = wst.Range("A:A").Find('NO', LookAt=1).Row
    trowf = get_last_row_from_column(wst, 'A')

    tr = trows
    for sr in tqdm(range(srows, srowf + 1), mininterval= 1):
        if int(wss.Cells(sr, 7).Value) == 0 :
            continue

        tr = tr + 1
        wst.Cells(tr, 1).Value = tr - trows - 1
        tmpstr = wss.Cells(sr, 1).Value
        wst.Cells(tr, 2).Value = tmpstr.replace('-', '')
        wst.Cells(tr, 3).Value = wss.Cells(sr, 2).Value
        wst.Cells(tr, 5).Value = wss.Cells(sr, 3).Value
        wst.Cells(tr, 6).Value = wss.Cells(sr, 4).Value
        wst.Cells(tr, 7).Value = wss.Cells(sr, 7).Value

    trim_blank_rows(wst)
    wst.Range("A1").Select()

    try:
        wbt.SaveAs(output_xlsx, FileFormat=51)
    except:
        print_message(0, 'save cancel ' + output_xlsx)
        wbt.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)

    wbs.Close()
    wbt.Close()
    excel_app.Application.Quit()

    print_message(9, 'finished')

    return output_xlsx

def get_column_after(aSheet, column, offset):
    ws = aSheet
    for item in ws.Range("{0}{1}:{0}{2}".format(column, offset, ws.get_last_row_from_column(column))).Value:
        print(item[0])
def get_last_row_from_column(aSheet, column):
    ws = aSheet
    return ws.Range("{0}{1}".format(column, ws.Rows.Count)).End(win32.constants.xlUp).Row

def trim_blank_rows(aSheet):
    print_message(2, 'starting ...')

    ws = aSheet
    rows = ws.Range("A:A").Find('NO', LookAt=1).Row
    rowf = get_last_row_from_column(ws, 'A')

    for tr in range(rowf, rows - 1, -1):

        lr = tr - rows - 1
        if (ws.Cells(tr, 2).Value is None) or (ws.Cells(tr, 2).Value == ''):
            if ws.Cells(tr, 1).Value % 30 == 0:
                if not ws.Cells(tr - 30 + 1, 2).Value is None:
                    break
            ws.Rows(tr).EntireRow.Delete()

    print_message(2, 'finished')

def print_message(asep, amesg):
    now = "[" + datetime.now().strftime("%m/%d/%Y-%H:%M:%S") +"]"
    if asep == 0:
        print('==========================================================')

    print(now, sys._getframe(1).f_code.co_name + "()", amesg, sep=':')

    if asep == 9:
        print('----------------------------------------------------------')

def zf_pdf_2_xls(input_pdf, output_xls):
    print_message(2, 'opening PDF ...')
    try:
        document = ap.Document(input_pdf)

        # 저장 옵션 생성 및 설정
        print_message(2, 'converting PDF -> xls')
        save_option = ap.ExcelSaveOptions()
        save_option.format = ap.ExcelSaveOptions.ExcelFormat.XML_SPREAD_SHEET2003

        # 파일을 MS Excel 형식으로 저장
        print_message(2, 'saving xls ... ' + output_xls)
        print_message(2, 'waiting (10 second) ... ')
        document.save(output_xls, save_option)
        print_message(2, 'saved          ' + output_xls)
    except:
        print_message(9, 'converting FAIL!')
        sys.exit(1)

import subprocess
def close_excel():
    """
    wb = xw.Book()

    wb.close()
    xw.App().quit()
    """

    # subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"])
    subprocess.call(["taskkill", "/im", "EXCEL.EXE"])
def main():
    close_excel()

    print_message(0, 'Starting ...')

    input_pdf = zf_FileLoad()

    if input_pdf == '':
        print_message(9, 'cancel')
        sys.exit(1)

    FileDir = os.path.dirname(input_pdf).replace("/", "\\")
    FileNam, FileExt = os.path.splitext(os.path.basename(input_pdf))
    output_xls = FileDir + "\\" + FileNam + ".xls"
    zf_pdf_2_xls(input_pdf, output_xls)
    """
    my_thread = threading.Thread(target =zf_pdf_2_xls, args=(input_pdf, output_xls,))
    my_thread.start()

    while my_thread.is_alive():
        print(".", end='')
        time.sleep(0.001)
    print('')
    """
    # 파일을 MS Excel 형식 변경 ( xls --> xlsx )
    print_message(2, 'converting xls -> xlsx')
    input_xlsx = output_xls
    output_xlsx = zf_Xls2Xlsx(input_xlsx)

    # 파일 생성 - 자재요청서
    print_message(2, 'creating MR')
    input_xlsx = output_xlsx
    output_xlsx = zf_MR_create(input_xlsx)

    print_message(9, 'finshed')

if __name__ == "__main__":
    main()
