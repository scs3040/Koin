import os
import sys
from datetime import datetime, timedelta
from tkinter import filedialog
import win32com.client as win32

def zs_write_head(a_wss, a_wst):
    zs_print_message(2, f'starting... ')

    wss = a_wss
    wst = a_wst

    #print(datetime.now().strftime('%m%d'))
    wst.Name = (datetime.today() + timedelta(days=1)).strftime('%m%d')

    sval = (datetime.today()+ timedelta(days=1)).strftime('%Y-%m-%d')
    zs_xl_put_string(wst, 'A:ZZ', '{반출일}', '0,0', sval)

    sval = zf_xl_get_string(wss, 'A:ZZ', '청구부서', '0,1')
    zs_xl_put_string(wst, 'A:ZZ', '{반출팀}', '0,0', sval)

    sval = zf_xl_get_string(wss, 'A:ZZ', '청구담당', '0,2')
    zs_xl_put_string(wst, 'A:ZZ', '{담당자}', '0,0', sval + ' 감독님')

    sval = zf_xl_get_string(wss, 'A:ZZ', '공사번호', '0,0')
    sval = sval.split(':')[1].strip()
    zs_xl_put_string(wst, 'A:ZZ', '{공사번호}', '0,0', sval)

    sval = zf_xl_get_string(wss, 'A:ZZ', '공사번호', '0,1')
    zs_xl_put_string(wst, 'A:ZZ', '{공사명}', '0,0', sval)

    zs_print_message(2, f'finished... ')


def zf_create_carryout(input_xlsx, output_xlsx):
    zs_print_message(0, f'starting ......')

    cur_dir = os.getcwd()
    #par_dir = os.path.abspath(os.path.join(cur_dir, os.pardir))
    file_tmpl = cur_dir + "\\_Tmpl\\_tmpl_반출요청서.xlsx"

    #file_dir = os.path.dirname(input_xlsx).replace("/", "\\")
    #file_name, file_ext = os.path.splitext(os.path.basename(input_xlsx))
    #output_xlsx = file_dir + "\\" + file_name + "_반출요청서.xlsx"
    try:
        os.remove(output_xlsx)
    except:
        zs_print_message(2, f'file not found  {output_xlsx}')
        zs_print_message(2, f'open .......... {output_xlsx}')

    excel_app = win32.gencache.EnsureDispatch("Excel.Application")

    try:
        wbs = excel_app.Workbooks.Open(input_xlsx)
        zs_print_message(2, f'open .......... {output_xlsx}')

    except:
        wbs.Close(SaveChanges=False)
        excel_app.Application.Quit()

        zs_print_message(2, f'open Fail .....')

        sys.exit(1)

    try:
        wbt = excel_app.Workbooks.Open(file_tmpl)
        zs_print_message(2, f'open Tmpl ..... {file_tmpl}')
    except:
        wbs.Close(SaveChanges=False)
        wbt.Close(SaveChanges=False)
        excel_app.Application.Quit()
        zs_print_message(2, f'open Fai l.....')
        sys.exit(1)

    zs_write_head(wbs.Sheets(1), wbt.Sheets(1))

    zs_print_message(2, f'creating ...... {output_xlsx}')

    wss = wbs.Sheets(1)
    srows = wss.Range('A:A').Find('품번', LookAt=1).Row + 1
    srowf = zf_get_last_row_from_column(wss, 'A')

    wst = wbt.Sheets(1)
    trows = wst.Range("A:B").Find('NO', LookAt=1).Row
    trowf = zf_get_last_row_from_column(wst, 'A') - 1

    trow = trows
    for srow in range(srows, srowf + 1):
        if int(wss.Cells(srow, 7).Value) == 0 :
            continue

        trow = trow + 1
        wst.Cells(trow, 1).Value = "=row()-11"
        tmp_str = wss.Cells(srow, 1).Value
        wst.Cells(trow, 3).Value = tmp_str.replace('-', '')
        wst.Cells(trow, 7).Value = wss.Cells(srow, 2).Value
        wst.Cells(trow, 13).Value = wss.Cells(srow, 3).Value
        wst.Cells(trow, 22).Value = wss.Cells(srow, 4).Value
        wst.Cells(trow, 25).Value = wss.Cells(srow, 7).Value

    #zs_trim_blank_rows(wst)
    wst.Range("A1").Select()

    try:
        excel_app.DisplayAlerts = False
        wbt.SaveAs(output_xlsx, FileFormat=51)
        excel_app.DisplayAlerts = True
        zs_print_message(2, f'saved ......... {output_xlsx}')

    except:
        wbt.Close(False)
        excel_app.Application.Quit()
        zs_print_message(0, f'save cancel ..')
        sys.exit(1)

    wbs.Close(False)
    wbt.Close()
    excel_app.Application.Quit()

    zs_print_message(9, 'finished')

    return output_xlsx

def zf_load_file():
    zs_print_message(2, 'select file ...')
    filename = filedialog.askopenfilename(initialdir="./", title="Select file",
                                          filetypes=(("Excel files", "*.xlsx"),
                                                     ("all files", "*.*")))
    zs_print_message(2, 'selected ' + filename)
    return filename


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


def zs_get_column_after(a_sheet, column, offset):
    ws = a_sheet
    for item in ws.Range("{0}{1}:{0}{2}".format(column, offset, ws.get_last_row_from_column(column))).Value:
        print(item[0])


def zf_get_last_row_from_column(aSheet, column):
    ws = aSheet
    return ws.Range("{0}{1}".format(column, ws.Rows.Count)).End(win32.constants.xlUp).Row


def zs_print_message(a_sep, a_mesg):
    now = '[' + datetime.now().strftime('%m/%d/%Y-%H:%M:%S') +']'
    if a_sep == 0:
        print('==========================================================')

    print(now, sys._getframe(1).f_code.co_name + "()", a_mesg, sep=':')

    if a_sep == 9:
        print('----------------------------------------------------------')


def zf_xl_get_string(a_ws, a_rng,  a_name, a_offset):
    frng = a_ws.Range(a_rng).Find(a_name, LookAt=2)
    ofrow, ofcol = a_offset.split(',')

    fcell = frng.GetOffset(int(ofrow), int(ofcol))
    return fcell.Value


def zs_xl_put_string(a_ws, a_rng,  a_name, a_offset, a_value):
    frng = a_ws.Range(a_rng).Find(a_name, LookAt=2)
    ofrow, ofcol = a_offset.split(',')

    fcell = frng.GetOffset(int(ofrow), int(ofcol))
    #frng.Cells(frng.Row + ofrow, frng.Column + ofcol).value = a_value
    fcell.Value = a_value


import argparse, sys
parser = argparse.ArgumentParser()
parser.add_argument('-input', help=' : Please set the input_xlsx(mast)')
parser.add_argument('-output', help=' : Please set the output_xlsx')

args = parser.parse_args()

def main(argv, args):

    zs_print_message(0, f'Starting ......')

    zs_print_message(2, f'argv : {argv}')
    zs_print_message(2, f'args : {args}')
    #-------------------------------------------------
    if args.input is None:
        input_xlsx = zf_load_file()
    else:
        input_xlsx = args.input

    if input_xlsx == '' or input_xlsx is None:
        zs_print_message(9, f'cancel.........')
        sys.exit(1)
    #-------------------------------------------------
    if args.output is None:
        file_dir = os.path.dirname(input_xlsx).replace("/", "\\")
        file_name, file_ext = os.path.splitext(os.path.basename(input_xlsx))
        output_xlsx = file_dir + "\\" + file_name + "_반출요청서.xlsx"
    else :
        output_xlsx = args.output

    #-------------------------------------------------
    list_wb = list()
    list_wb.append(file_name)
    list_wb.append('Tmpl_반출요청서.xlsx')

    zf_close_all_wb(list_wb)

    #-------------------------------------------------
    # 파일 생성 - 반출요청서
    zs_print_message(2, 'creating ....... _반출요청서')
    output_xlsx = zf_create_carryout(input_xlsx, output_xlsx)
    zs_print_message(2, f'create success! _반출요청서')

    #-------------------------------------------------
    zs_print_message(9, 'finshed ........')


#=====================================================
if __name__ == "__main__":
    argv = sys.argv
    main(argv, args)
