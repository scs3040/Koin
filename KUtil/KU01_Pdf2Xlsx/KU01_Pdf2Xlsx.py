import os
from tkinter import *
from tkinter import filedialog

#import jpype
import aspose.pdf as ap
import asposecells

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


""" def zf_Xls2Xlsx(input_xls) :

    jpype.startJVM()

    from asposecells.api import Workbook

    FileDir = os.path.dirname(input_pdf)
    FileNam, FileExt = os.path.splitext(os.path.basename(input_pdf))

    workbook = Workbook(input_xls)

    output_xlsx = FileDir + "\\" + FileNam + ".xlsx"
    workbook.save(output_xlsx)

    jpype.shutdownJVM()

    return output_xlsx """


def zf_Xls2Xlsx(input_xls):
    FileDir = os.path.dirname(input_xls)
    FileNam, FileExt = os.path.splitext(os.path.basename(input_xls))
    output_xlsx = FileDir + "\\" + FileNam + "_new.xlsx"

    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        wb = excel_app.Workbooks.Open(input_xls)
    except:
        print
        "Failed to open spreadsheet " + FileNam
        sys.exit(1)

    wscnt = wb.Sheets.Count

    print(wscnt)

#    for (int i = 1; i <= wscnt; i++):
    wst = wb.Sheets(1)
    for i in range(2,wscnt+1):

        wss = wb.Sheets(i)
        #wss.Rows(1).EntireRow.Delete()

        lastrow = wst.UsedRange.Rows.Count

        for i in range(lastrow, 0, -1):
            str1 = wss.Cells(i, 1).Value
            print(str1)
            if str1 is None:
                wss.Rows(i).EntireRow.Delete()
            else:
                str2 = str1[0:2]
                print(str2)
                if not str2.isnumeric():
                    wss.Rows(i).EntireRow.Delete()

        wss.UsedRange.Copy()

        lastrow = wst.UsedRange.Rows.Count-2
        wst.Cells(lastrow, 1).Select()
        wst.Paste()
        wst.Cells(1, 1).Value = ''

    wb.SaveAs(output_xlsx, FileFormat=51)
    wb.Close()

    excel_app.Application.Quit()

    return output_xlsx


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

output_xlsx = zf_Xls2Xlsx(output_xls)
