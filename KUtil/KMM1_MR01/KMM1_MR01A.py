import os
import sys
#from tkinter import *
from tkinter import filedialog

import aspose.pdf as ap
#import asposecells

import win32com.client as win32


def zf_FileLoad():
    filename = filedialog.askopenfilename(initialdir="./", title="Select file",
                                          filetypes=(("PDF files", "*.pdf"),
                                                     ("all files", "*.*")))
    # print(filename)
    return filename


def zf_FileSave():
    filename = filedialog.asksaveasfilename(initialdir="./", title="Select file",
                                            filetypes=(("XLS files", "*.xls"),
                                                       ("all files", "*.*")))
    # print(filename)
    return filename


def zf_Xls2Xlsx(input_xls):
    FileDir = os.path.dirname(input_xls).replace("/", "\\")
    FileNam, FileExt = os.path.splitext(os.path.basename(input_xls))
    output_xlsx = FileDir + "\\" + FileNam + "_1.xlsx"

    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        wb = excel_app.Workbooks.Open(input_xls)
    except:
        print("Failed to open spreadsheet " + FileNam)
        sys.exit(1)

    zs_SheetMerge(wb)

    wb.SaveAs(output_xlsx, FileFormat=51)
    wb.Close()

    excel_app.Application.Quit()

    return output_xlsx


def zs_SheetMerge(aWorkbook):
    wb = aWorkbook

    wsCnt = wb.Sheets.Count
    wst = wb.Sheets(1)
    for i in range(2, wsCnt + 1):

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

        zs_SheetStyle(wst, "A:Z")


def zs_SheetStyle(aWorksheet, aRange ):
    ws = aWorksheet
    rng = aRange

    ws.Range(aRange).Font.Size = 10


def zf_MR_create(input_xlsx):
    FileDir = os.path.dirname(input_xlsx).replace("/", "\\")
    FileNam, FileExt = os.path.splitext(os.path.basename(input_xlsx))
    CurrDir = os.path.dirname(os.path.realpath(__file__))
    FileTmpl = CurrDir + "\\_Tmpl\\tmpl_자재요청.xlsx"
    output_xlsx = FileDir + "\\" + FileNam + "_자재요청서.xlsx"

    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        wbs = excel_app.Workbooks.Open(input_xlsx)
        wbt = excel_app.Workbooks.Open(FileTmpl)

    except:
        wbs.Close()
        wbt.Close()
        sys.exit(1)

    wss = wbs.Sheets("page 1")
    srows = wss.Range("A:A").Find('품번', LookAt=1).Row +1
    srowf = get_last_row_from_column(wss, 'A')

    wst = wbt.Sheets(1)
    trows = wst.Range("A:A").Find('NO', LookAt=1).Row
    trowf = get_last_row_from_column(wst, 'A')

    tr = trows
    for sr in range(srows, srowf + 1):
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

    try:
        wbt.SaveAs(output_xlsx, FileFormat=51)
    except:
        print("err")

    wbs.Close()
    wbt.Close()

    excel_app.Application.Quit()

    return output_xlsx

def get_column_after(aSheet, column, offset):
    ws = aSheet
    for item in ws.Range("{0}{1}:{0}{2}".format(column, offset, ws.get_last_row_from_column(column))).Value:
        print(item[0])
def get_last_row_from_column(aSheet, column):
    ws = aSheet
    return ws.Range("{0}{1}".format(column, ws.Rows.Count)).End(win32.constants.xlUp).Row

def trim_blank_rows(aSheet):
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

def main():
    input_pdf = zf_FileLoad()
    FileDir = os.path.dirname(input_pdf).replace("/", "\\")
    FileNam, FileExt = os.path.splitext(os.path.basename(input_pdf))
    output_xls = FileDir + "\\" + FileNam + ".xls"

    # PDF 파일 열기
    document = ap.Document(input_pdf)

    # 저장 옵션 생성 및 설정
    save_option = ap.ExcelSaveOptions()
    save_option.format = ap.ExcelSaveOptions.ExcelFormat.XML_SPREAD_SHEET2003

    # 파일을 MS Excel 형식으로 저장
    document.save(output_xls, save_option)

    # 파일을 MS Excel 형식 변경 ( xls --> xlsx )
    input_xlsx = output_xls
    output_xlsx = zf_Xls2Xlsx(input_xlsx)
    print(output_xlsx)
    # 파일 생성 - 자재요청서
    input_xlsx = output_xlsx
    output_xlsx = zf_MR_create(input_xlsx)


if __name__ == "__main__":
    main()
