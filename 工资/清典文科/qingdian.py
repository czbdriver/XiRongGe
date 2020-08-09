# -*- coding: utf-8 -*-
import xlrd
import xlutils
import xlwt
#from datetime import date,datetime
from tkinter import *
import tkinter.filedialog
import tkinter.messagebox
import sys
from tkinter import ttk
from xlutils.copy import copy
from xlutils.filter import process, XLRDReader, XLWTWriter
#from _pydecimal import Decimal, Context, ROUND_HALF_UP

import decimal
decimal.getcontext().rounding = "ROUND_HALF_UP"

root_main = Tk()

root_input = tkinter.LabelFrame(width=680, height=180,text='输入选项')
lbl_input_year_month = Label(root_input, text="请输入年月(格式:YYYYMM, 例如:201905):")
txt_input_year_month = Entry(root_input, bd =1)

var_show_input_last_file = StringVar()
var_show_input_last_file.set('未选择计算文件')
lbl_show_input_last_file = Label(root_input, textvariable=var_show_input_last_file, justify='left', fg='green')


def My_decimal(input):
    if (input == ''):
        input = 0
    return decimal.Decimal(str(input)).quantize(decimal.Decimal("0.00"))
    

def select_input_last_file():
    return

def start_app():
    root_main.title('清典工资条')
    root_main.geometry('700x205')

    root_input.place(x=10,y=20)
    lbl_input_year_month.place(x=30, y=10)
    txt_input_year_month.place(x=350, y=10)


    btn_input_last_file = Button(root_input, text="选择计算文件", width=75, command=select_input_last_file)
    btn_input_last_file.place(x=30, y=50)
    lbl_show_input_last_file.place(x=30, y=95)
    
    root_main.mainloop()


if __name__ == '__main__':
    start_app()
