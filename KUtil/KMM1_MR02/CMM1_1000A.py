import os
import sys
from datetime import datetime, timedelta

def zs_print_message(a_sep, a_mesg):
    now = "[" + datetime.now().strftime("%Y/%m/%d %H:%M:%S") +"]"
    if a_sep == 0:
        print('==========================================================')

    print(now, sys._getframe(1).f_code.co_name + "()", a_mesg, sep=':')

    if a_sep == 9:
        print('----------------------------------------------------------')