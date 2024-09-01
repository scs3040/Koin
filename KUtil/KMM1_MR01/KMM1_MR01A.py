import os
import sys
from datetime import datetime
from tkinter import filedialog

import aspose.pdf as ap

import win32com.client as win32
from tqdm import tqdm
import xlwings as xw

#import threading
#import time

def zf_load_file():
    zs_print_message(2, 'select file ...')
    filename = filedialog.askopenfilename(initialdir="./", title="Select file",
                                          filetypes=(("PDF files", "*.pdf"),
                                                     ("all files", "*.*")))
    zs_print_message(2, 'selected ' + filename)
    return filename


def zf_save_file():
    zs_print_message(2, 'select file ... ')
    filename = filedialog.asksaveasfilename(initialdir="./", title="Select file",
                                            filetypes=(("XLS files", "*.xls"),
                                                       ("all files", "*.*")))
    # print(filename)
    zs_print_message(2, 'Saved ' + filename)
    return filename


def zf_xls_2_xlsx(input_xls):
    zs_print_message(0, 'Starting ...')

    file_dir = os.path.dirname(input_xls).replace("/", "\\")
    file_name, file_ext = os.path.splitext(os.path.basename(input_xls))
    output_xlsx = file_dir + "\\" + file_name + "_1.xlsx"

    zs_print_message(2, 'open PDF ' + input_xls)
    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        wb = excel_app.Workbooks.Open(input_xls, Notify=False)
    except:
        wb.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)

    zs_merge_sheet(wb)

    zs_print_message(2, 'saving ... ' + output_xlsx)
    try:
        excel_app.DisplayAlerts = False
        wb.SaveAs(output_xlsx, FileFormat=51)
        excel_app.DisplayAlerts = True
    except:
        print_message(9, 'save cancel ')
        wb.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)

    zs_print_message(2, 'saved      ' + output_xlsx)
    wb.Close()

    excel_app.Application.Quit()

    zs_print_message(9, 'Finished')

    return output_xlsx


def zs_merge_sheet(a_workbook):
    wb = a_workbook
    zs_print_message(2, 'starting ...')

    wscnt = wb.Sheets.Count
    wst = wb.Sheets(1)
    for i in tqdm(range(2, wscnt + 1), mininterval=1):

        wss = wb.Sheets(i)

        lastrow = wst.UsedRange.Rows.Count

        for i in range(lastrow, 0, -1):
            tstr1 = wss.Cells(i, 1).Value
            if tstr1 is None:
                wss.Rows(i).EntireRow.Delete()
            else:
                tstr2 = tstr1[0:2]
                if not tstr2.isnumeric():
                    wss.Rows(i).EntireRow.Delete()

        wss.UsedRange.Copy()

        lastrow = wst.UsedRange.Rows.Count - 2
        wst.Cells(lastrow, 1).Select()
        wst.Paste()
        wst.Cells(1, 1).Value = ''

    zs_print_message(2, 'finished')
    zs_set_sheet_style(wst, "A:Z")
def zs_set_sheet_style(a_worksheet, a_range ):
    zs_print_message(2, 'starting .. ')

    lws = a_worksheet
    lrng = a_range

    lws.Range(a_range).Font.Size = 10

    zs_print_message(2, 'finished')


def zf_create_mr(input_xlsx):
    zs_print_message(0, 'starting ...')

    file_dir = os.path.dirname(input_xlsx).replace("/", "\\")
    file_name, file_ext = os.path.splitext(os.path.basename(input_xlsx))

    curr_dir = os.getcwd()
    file_tmpl = curr_dir + "\\_Tmpl\\tmpl_자재요청.xlsx"

    output_xlsx = file_dir + "\\" + file_name + "_자재요청서.xlsx"

    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        wbs = excel_app.Workbooks.Open(input_xlsx)
    except:
        wbs.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)

    try:
        wbt = excel_app.Workbooks.Open(file_tmpl)
    except:
        wbt.Close(SaveChanges=False)
        excel_app.Application.Quit()
        sys.exit(1)

    wss = wbs.Sheets(1)
    srows = wss.Range("A:A").Find('품번', LookAt=1).Row + 1
    srowf = zf_get_last_row_from_column(wss, 'A')

    wst = wbt.Sheets(1)
    trows = wst.Range("A:A").Find('NO', LookAt=1).Row
    trowf = zf_get_last_row_from_column(wst, 'A')

    trow = trows
    for srow in tqdm(range(srows, srowf + 1), mininterval= 1):
        if int(wss.Cells(srow, 7).Value) == 0 :
            continue

        trow = trow + 1
        wst.Cells(trow, 1).Value = trow - trows - 1
        tmp_str = wss.Cells(srow, 1).Value
        wst.Cells(trow, 2).Value = tmp_str.replace('-', '')
        wst.Cells(trow, 3).Value = wss.Cells(srow, 2).Value
        wst.Cells(trow, 5).Value = wss.Cells(srow, 3).Value
        wst.Cells(trow, 6).Value = wss.Cells(srow, 4).Value
        wst.Cells(trow, 7).Value = wss.Cells(srow, 7).Value

    zs_trim_blank_rows(wst)
    wst.Range("A1").Select()

    try:
        excel_app.DisplayAlerts = False
        wbt.SaveAs(output_xlsx, FileFormat=51)
        excel_app.DisplayAlerts = True
    except:
        zs_print_message(0, 'save cancel ' + output_xlsx)
        wbt.Close(False)
        excel_app.Application.Quit()
        sys.exit(1)

    wbs.Close()
    wbt.Close()
    excel_app.Application.Quit()

    zs_print_message(9, 'finished')

    return output_xlsx

def zs_get_column_after(a_sheet, column, offset):
    ws = a_sheet
    for item in ws.Range("{0}{1}:{0}{2}".format(column, offset, ws.get_last_row_from_column(column))).Value:
        print(item[0])
def zf_get_last_row_from_column(aSheet, column):
    ws = aSheet
    return ws.Range("{0}{1}".format(column, ws.Rows.Count)).End(win32.constants.xlUp).Row

def zs_trim_blank_rows(aSheet):
    zs_print_message(2, 'starting ...')

    ws = aSheet
    rows = ws.Range("A:A").Find('NO', LookAt=1).Row
    rowf = zf_get_last_row_from_column(ws, 'A')

    for trow in range(rowf, rows - 1, -1):

        lrow = trow - rows - 1
        if (ws.Cells(trow, 2).Value is None) or (ws.Cells(trow, 2).Value == ''):
            if ws.Cells(trow, 1).Value % 30 == 0:
                if not ws.Cells(trow - 30 + 1, 2).Value is None:
                    break
            ws.Rows(trow).EntireRow.Delete()

    zs_print_message(2, 'finished')

def zs_print_message(a_sep, a_mesg):
    now = "[" + datetime.now().strftime("%m/%d/%Y-%H:%M:%S") +"]"
    if a_sep == 0:
        print('==========================================================')

    print(now, sys._getframe(1).f_code.co_name + "()", a_mesg, sep=':')

    if a_sep == 9:
        print('----------------------------------------------------------')

def zf_pdf_2_xls(input_pdf, output_xls):
    zs_print_message(2, 'opening PDF ...')
    try:
        document = ap.Document(input_pdf)

        # 저장 옵션 생성 및 설정
        zs_print_message(2, 'converting PDF -> xls')
        save_option = ap.ExcelSaveOptions()
        save_option.format = ap.ExcelSaveOptions.ExcelFormat.XML_SPREAD_SHEET2003

        # 파일을 MS Excel 형식으로 저장
        zs_print_message(2, 'saving xls ... ' + output_xls)
        zs_print_message(2, 'waiting (10 second) ... ')

        document.save(output_xls, save_option)

        zs_print_message(2, 'saved          ' + output_xls)
    except:
        zs_print_message(9, 'converting FAIL!')
        sys.exit(1)

import subprocess
def zf_close_all_wb(alist_wbname):

    llist = alist_wbname

    com_app = win32.dynamic.Dispatch('Excel.Application')
    com_wbs = com_app.Workbooks
    list_wb_names = [wb.Name for wb in com_wbs]

    zs_print_message(2, 'closing ' + str(list_wb_names))

    lcnt = com_wbs.Count
    for i in reversed(range(lcnt)):
        wb_name = com_wbs[i].Name
        if wb_name == 'tmpl_자재요청.xlsx':
            com_wbs[i].Close(SaveChanges=False)

        for wbnam in alist_wbname:
            if wbnam in wb_name:
                zs_print_message(2, 'closed ' + wb_name)
                com_wbs[i].Close(SaveChanges=False)
    com_app.Quit()

    # subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"])
    #subprocess.call(["taskkill", "/im", "EXCEL.EXE"])
def main():

    zs_print_message(0, 'Starting ...')

    input_pdf = zf_load_file()

    if input_pdf == '':
        zs_print_message(9, 'cancel')
        sys.exit(1)

    file_dir = os.path.dirname(input_pdf).replace("/", "\\")
    file_name, file_ext = os.path.splitext(os.path.basename(input_pdf))
    output_xls = file_dir + "\\" + file_name + ".xls"

    list_wb = list()
    list_wb.append(file_name)
    list_wb.append('Tmpl_자재요청서.xlsx')

    zf_close_all_wb(list_wb)

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
    zs_print_message(2, 'converting xls -> xlsx')
    input_xlsx = output_xls
    output_xlsx = zf_xls_2_xlsx(input_xlsx)

    # 파일 생성 - 자재요청서
    zs_print_message(2, 'creating MR')
    input_xlsx = output_xlsx
    output_xlsx = zf_create_mr(input_xlsx)

    zs_print_message(9, 'finshed')

if __name__ == "__main__":
    main()
