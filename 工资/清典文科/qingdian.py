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
root_str = tkinter.LabelFrame(width=680, height=160,text='字符对应')
lbl_pay_total = Label(root_str, text="本期收入")
txt_pay_total = Entry(root_str, width=15, bd =1)
txt_pay_total.insert(0, '应发合计')

lbl_yanglao = Label(root_str, text="本期基本养老保险费")
txt_yanglao = Entry(root_str, width=15, bd =1)
txt_yanglao.insert(0, '养老个人')

lbl_yiliao = Label(root_str, text="本期基本医疗保险费")
txt_yiliao = Entry(root_str, width=15, bd =1)
txt_yiliao.insert(0, '医疗个人')

lbl_shiye = Label(root_str, text="本期失业保险费")
txt_shiye = Entry(root_str, width=15, bd =1)
txt_shiye.insert(0, '失业个人')

lbl_gongjijin = Label(root_str, text="本期住房公积金")
txt_gongjijin = Entry(root_str, width=15, bd =1)
txt_gongjijin.insert(0, '公积金个人')

lbl_this_start = Label(root_str, text="本月工资文件行数")
txt_this_start = Entry(root_str, width=5, bd =1, fg='blue')
txt_this_start.insert(0, '1')

lbl_this_end = Label(root_str, text=" - ", fg='blue')
txt_this_end = Entry(root_str, width=5, bd =1, fg='blue')
txt_this_end.insert(0, '205')

root_input = tkinter.LabelFrame(width=680, height=315,text='输入选项')
lbl_input_year_month = Label(root_input, text="请输入年月(格式:YYYYMM, 例如:201905):")
txt_input_year_month = Entry(root_input, bd =1)

var_show_input_last_file = StringVar()
var_show_input_last_file.set('未选择上月税款计算文件')
lbl_show_input_last_file = Label(root_input, textvariable=var_show_input_last_file, justify='left', fg='green')

var_show_input_this_file = StringVar()
var_show_input_this_file.set('未选择本月工资明细文件')
lbl_show_input_this_file = Label(root_input, textvariable=var_show_input_this_file, justify='left', fg='blue')

root_output = tkinter.LabelFrame(width=680, height=140,text='税款计算生成')
var_show_root_output_file = StringVar()
var_show_root_output_file.set('生成文件路径：')
lbl_show_root_output_file = Label(root_output, textvariable=var_show_root_output_file, justify='left', fg='purple')

g_input_deducation_file_path = ''
var_show_input_deducation_file = StringVar()
var_show_input_deducation_file.set('未选择专项扣除文件')
lbl_show_input_deducation_file = Label(root_input, textvariable=var_show_input_deducation_file, justify='left', fg='blue')
g_list_deduction = []

title_gen = ['工号','姓名','证照类型','证照号码','税款所属期起',
             '税款所属期止','所得项目','本期收入','本期费用','本期免税收入',
             '本期基本养老保险费','本期基本医疗保险费','本期失业保险费','本期住房公积金','本期企业(职业)年金',
             '本期商业健康保险费','本期税延养老保险费','本期其他扣除(其他)','累计收入额','累计减除费用',
             '累计专项扣除','累计子女教育支出扣除','累计赡养老人支出扣除','累计继续教育支出扣除','累计住房贷款利息支出扣除',
             '累计住房租金支出扣除','累计其他扣除','累计准予扣除的捐赠','累计应纳税所得额','税率',
             '速算扣除数','累计应纳税额','累计减免税额','累计应扣缴税额','累计已预缴税额',
             '累计应补(退)税额','备注']

g_input_year_month = ''
btn_input_last_file = ''
btn_input_this_file = ''
btn_input_deduction_file = ''

g_input_last_file_path = '' 
g_input_this_file_path = ''
g_input_this_file_dir = ''
g_output_file_path = ''

g_str_this_total = '应发合计'
g_str_this_yanglao = '养老个人'
g_str_this_shiye = '失业个人'
g_str_this_yiliao = '医疗个人'
g_str_this_gongjijin = '公积金个人'
g_str_this_fangbu = '房补'
g_str_this_duzi = '独子'
g_str_this_fubu = '副补'
g_str_this_tax_arrears = '个税补交'

g_str_this_gonghao = '工号'
g_str_this_name = '姓名'

g_person_list = []        
g_list_person_last_info = []
g_list_person_this_info = []

g_input_this_file_start_index = 1
g_input_this_file_end_index = 181

# 子女教育
class adjust_err_info:
     def __init__(self, name, value):
         self.name = name         # name
         self.value = value       # deduct value

# 子女教育
class children_deduction_info:
     def __init__(self, name, value):
         self.name = name         # name
         self.value = value       # deduct value

# 赡养老人
class parents_deduction_info:
     def __init__(self, name, value):
         self.name = name         # name
         self.value = value       # deduct value

# 继续教育         
class education_deduction_info:
     def __init__(self, name, value):
         self.name = name         # name
         self.value = value       # deduct value

#  贷款利息
class loans_deduction_info:
     def __init__(self, name, value):
         self.name = name         # name
         self.value = value       # deduct value

#  租房租金
class rent_deduction_info:
     def __init__(self, name, value):
         self.name = name         # name
         self.value = value       # deduct value

class deduction_info:
    def __init__(self):
        self.name = ''
        self.children = 0
        self.parents = 0
        self.education = 0
        self.loans = 0
        self.rent = 0

class model_person_info:
    def __init__(self):
        self.item1 = ""              # 工号
        self.item2 = ""              # 姓名
        self.item3 = ""              # 证照类型
        self.item4 = ""              # 证照号码
        self.item5 = ""              # 税款所属期起
        self.item6 = ""              # 税款所属期止
        self.item7 = ""              # 所得项目
        self.item8 = ""              # 本期收入
        self.item9 = ""              # 本期费用
        self.item10 = ""              # 本期免税收入
        self.item11 = ""              # 本期基本养老保险费
        self.item12 = ""              # 本期基本医疗保险费
        self.item13 = ""              # 本期失业保险费
        self.item14 = ""              # 本期住房公积金
        self.item15 = ""              # 本期企业(职业)年金
        self.item16 = ""              # 本期商业健康保险费
        self.item17 = ""              # 本期税延养老保险费
        self.item18 = ""              # 本期其他扣除(其他)
        self.item19 = ""              # 累计收入额
        self.item20 = ""              # 累计减除费用
        self.item21 = ""              # 累计专项扣除
        self.item22 = ""              # 累计子女教育支出扣除
        self.item23 = ""              # 累计赡养老人支出扣除
        self.item24 = ""              # 累计继续教育支出扣除
        self.item25 = ""              # 累计住房贷款利息支出扣除
        self.item26 = ""              # 累计住房租金支出扣除
        self.item27 = ""              # 累计其他扣除
        self.item28 = ""              # 累计准予扣除的捐赠
        self.item29 = ""              # 累计应纳税所得额
        self.item30 = ""              # 税率
        self.item31 = ""              # 速算扣除数
        self.item32 = ""              # 累计应纳税额
        self.item33 = ""              # 累计减免税额
        self.item34 = ""              # 累计应扣缴税额
        self.item35 = ""              # 累计已预缴税额
        self.item36 = ""              # 累计应补(退)税额
        self.item37 = ""              # 备注

def My_decimal(input):
    if (input == ''):
        input = 0
    return decimal.Decimal(str(input)).quantize(decimal.Decimal("0.00"))
    
def has_str(str, list):
    for index in range(len(list)):
        if str == list[index]:
            return 1
    return 0

def get_all_deducation():
    g_input_deducation_file_path = tkinter.filedialog.askopenfilename()
    if g_input_deducation_file_path != '':
        print('the selected input this file is:'+g_input_deducation_file_path)
        var_show_input_deducation_file.set('选择文件: ' + g_input_deducation_file_path)
    else:
        print('no file selected')
        return
    
    workbook = xlrd.open_workbook(g_input_deducation_file_path)
    sheet1 = workbook.sheet_by_index(0)

    row_title = sheet1.row_values(0)

    g_list_deduction.clear()
    for index in range(sheet1.nrows):
        if (index > 0):
            tmp = deduction_info()
            list_rows = (sheet1.row_values(index))
            for ii in range(len(row_title)):
                if row_title[ii] == "姓名":
                    tmp.name = list_rows[ii]
                elif row_title[ii] == "子女教育支出扣除":
                    tmp.children = list_rows[ii]
                elif row_title[ii] == "继续教育支出扣除":
                    tmp.education = list_rows[ii]    
                elif row_title[ii] == "住房贷款利息支出扣除":
                    tmp.loans = list_rows[ii]
                elif row_title[ii] == "住房租金支出扣除":
                    tmp.rent = list_rows[ii]
                elif row_title[ii] == "赡养老人支出扣除":
                    tmp.parents = list_rows[ii]
            g_list_deduction.append(tmp)

def get_deduction_value(name, key):
    global g_list_deduction

    for index in range(len(g_list_deduction)):
        if name == g_list_deduction[index].name:
            if key == "子女教育支出扣除":
                return g_list_deduction[index].children
            elif key == "继续教育支出扣除":
                return g_list_deduction[index].education
            elif key == "住房贷款利息支出扣除":
                return g_list_deduction[index].loans
            elif key == "住房租金支出扣除":
                return g_list_deduction[index].rent
            elif key == "赡养老人支出扣除":
                return g_list_deduction[index].parents

    return 0

def select_input_this_file():
    global g_list_person_this_info
    global g_str_this_total
    global g_input_this_file_path
    global g_input_this_file_start_index
    global g_input_this_file_end_index
    global g_input_this_file_dir
    global g_str_this_tax_arrears
    
    if (check_warning_year_month()== 0):
        return
    
    if (check_warning_input_last_file()== 0):
        return
    get_input_str()
    g_input_this_file_path = tkinter.filedialog.askopenfilename()
    if g_input_this_file_path != '':
        print('the selected input this file is:'+g_input_this_file_path)
        var_show_input_this_file.set('选择文件: ' + g_input_this_file_path)
        
        tmp = str(g_input_this_file_path).split('/')
        
        g_input_this_file_dir = ''
        for index in range(len(tmp)-1):
            g_input_this_file_dir = g_input_this_file_dir + tmp[index] + '/'
        
    else:
        print('no file selected')
        return
    
    workbook = xlrd.open_workbook(g_input_this_file_path)
    sheet1 = workbook.sheet_by_index(0)
    
    g_list_person_this_info.clear()
    
    row_title = sheet1.row_values(0)

    if has_str(g_str_this_total, row_title) == 0:
        tkinter.messagebox.showwarning("警告", "工资明细文件不包含 "+g_str_this_total)
        txt_pay_total.focus_set()
        return
    if has_str(g_str_this_shiye, row_title) == 0:
        tkinter.messagebox.showwarning("警告", "工资明细文件不包含 "+g_str_this_shiye)
        txt_shiye.focus_set()
        return
    if has_str(g_str_this_yanglao, row_title) == 0:
        tkinter.messagebox.showwarning("警告", "工资明细文件不包含 "+g_str_this_yanglao)
        txt_yanglao.focus_set()
        return
    if has_str(g_str_this_yiliao, row_title) == 0:
        tkinter.messagebox.showwarning("警告", "工资明细文件不包含 "+g_str_this_yiliao)
        txt_yiliao.focus_set()
        return
    if has_str(g_str_this_gongjijin, row_title) == 0:
        tkinter.messagebox.showwarning("警告", "工资明细文件不包含 "+g_str_this_gongjijin)
        txt_gongjijin.focus_set()
        return

    for index in range(sheet1.nrows):
        if (index >=  g_input_this_file_start_index) and (index <  g_input_this_file_end_index):
            list_rows = (sheet1.row_values(index))
            tmp_item = model_person_info()
            tmp_item.item8 = My_decimal(0)
            tmp_item.item10 = My_decimal(0)
            for ii in range(len(row_title)):
                if (row_title[ii] == g_str_this_total):
                    tmp_item.item8 = tmp_item.item8 + My_decimal(list_rows[ii])
                elif (row_title[ii] == g_str_this_shiye):
                    tmp_item.item13 = My_decimal(list_rows[ii])
                elif (row_title[ii] == g_str_this_yanglao):
                    tmp_item.item11 = My_decimal(list_rows[ii])
                elif (row_title[ii] == g_str_this_yiliao):
                    tmp_item.item12 = My_decimal(list_rows[ii])
                elif (row_title[ii] == g_str_this_gongjijin):
                    tmp_item.item14 = My_decimal(list_rows[ii])
                elif (row_title[ii] == g_str_this_gonghao):
                    tmp_item.item1 = str(list_rows[ii])
                elif (row_title[ii] == g_str_this_name):
                    tmp_item.item2 = str(list_rows[ii])
            
            g_list_person_this_info.append(tmp_item)      
    return 0

def select_input_deduction_file():
    if (check_warning_year_month()== 0):
        return
    get_all_deducation()

def get_real_index(key, list_title):
    global title_gen

    index = int(key[4:])
    index = index -1

    for ii in range(len(list_title)):
        if list_title[ii] == title_gen[index]:
            return ii

    return -1

def select_input_last_file():
    global g_list_person_last_info
    global g_input_last_file_path
    
    if (check_warning_year_month()== 0):
        return

    g_input_last_file_path = tkinter.filedialog.askopenfilename()
    
    if g_input_last_file_path != '':
        print('the selected input last file is:'+g_input_last_file_path)
        var_show_input_last_file.set('选择文件: ' + g_input_last_file_path)
    else:
        print('no file selected')
        return
        
    workbook = xlrd.open_workbook(g_input_last_file_path)
    sheet1 = workbook.sheet_by_index(0)
    g_list_person_last_info.clear()

    list_title = sheet1.row_values(0)

    for index in range(sheet1.nrows):
        if (index > 0):
            list_rows = (sheet1.row_values(index))  
            tmp_item = model_person_info()

            for str_item in vars(tmp_item).keys():
                #i_tmp = int(str_item[4:])
                #i_tmp = i_tmp -1
                real_index = get_real_index(str_item, list_title)
                vars(tmp_item)[str_item] = list_rows[real_index]
            g_list_person_last_info.append(tmp_item)
    print("end")

def check_warning_year_month():
    global g_input_year_month
    
    g_input_year_month = txt_input_year_month.get()
    if not str(g_input_year_month).strip():
        tkinter.messagebox.showwarning("警告", "请先输入年月")
        txt_input_year_month.focus_set()
        return 0
    return 1

def check_warning_input_last_file():
    global g_input_last_file_path
    
    if not str(g_input_last_file_path).strip():
        tkinter.messagebox.showwarning("警告", "请先点击'选择上月税款计算文件'按钮选择上月个税计算文件")
        btn_input_last_file.focus_set()
        return 0
    return 1

def check_warning_input_this_file():
    global g_input_this_file_path
    
    if not str(g_input_this_file_path).strip():
        tkinter.messagebox.showwarning("警告", "请先点击'选择本月工资明细文件'按钮选择本月工资计算文件")
        btn_input_this_file.focus_set()
        return 0
    return 1

def get_last_info(name, key_index):
    global g_list_person_last_info
    
    for index in range(len(g_list_person_last_info)):
        if name == g_list_person_last_info[index].item2:
            str_item = 'item' + str(key_index)
            return vars(g_list_person_last_info[index])[str_item]
    return 0

def find_last_info(name):
    global g_list_person_last_info
    
    for index in range(len(g_list_person_last_info)):
        if name == g_list_person_last_info[index].item2:
            return 1
    return 0
    
def merge_person_info():
    global g_list_person_last_info
    global g_list_person_this_info
    global g_person_list
    global g_input_year_month
    
    #get_input_str()
    g_person_list.clear()
    for index in range(len(g_list_person_this_info)):
        tmp_item = model_person_info()
        # 工号1
        tmp_item.item1 = str(g_list_person_this_info[index].item1)
        # 姓名2
        tmp_item.item2 = str(g_list_person_this_info[index].item2)
        # 本期收入8
        tmp_item.item8 = My_decimal(g_list_person_this_info[index].item8)
        # 本期基本养老保险费11
        tmp_item.item11 = My_decimal(g_list_person_this_info[index].item11)
        # 本期基本医疗保险费12
        tmp_item.item12 = My_decimal(g_list_person_this_info[index].item12)
        # 本期失业保险费13
        tmp_item.item13 = My_decimal(g_list_person_this_info[index].item13)
        # 本期住房公积金14
        tmp_item.item14 = My_decimal(g_list_person_this_info[index].item14)
        # 本期免税收入10
        tmp_item.item10 = My_decimal(g_list_person_this_info[index].item10)
        
        # 税款所属期起5
        tmp_item.item5 = g_input_year_month + "01"
        # 税款所属期止6
        year = int(g_input_year_month[0:4])
        month = int(g_input_year_month[5:])
        end = 0
        if ((month == 1) or (month == 3)
            or (month == 5) or (month == 7)
            or (month == 8) or (month == 10)
            or (month == 12)):
            end = 31
        elif(month == 2):
            if (year%4==0) and (year%100 == 0):
                end = 28
            elif(year%4 == 0):
                end = 29
            else:
                end = 28
        else:
            end = 30
        tmp_item.item6 = g_input_year_month + str(end)
        
        # 所得项目7
        tmp_item.item7 = "正常工资薪金"
        # 本期费用9
        tmp_item.item9 = My_decimal(0)
        
        # 本期企业(职业)年金15
        tmp_item.item15 = My_decimal(0)
        # 本期商业健康保险费16
        tmp_item.item16 = My_decimal(0)
        # 本期税延养老保险费17
        tmp_item.item17 = My_decimal(0)
        # 本期其他扣除(其他)18
        tmp_item.item18 = My_decimal(0)
        
        # 累计收入额19
        if find_last_info(tmp_item.item2):
            tmp_item.item19 = (My_decimal(tmp_item.item8) + My_decimal(get_last_info(tmp_item.item2, 19)))
        else:
            tmp_item.item19 = (My_decimal(tmp_item.item8))
        # 累计减除费用20
        if find_last_info(tmp_item.item2):
            tmp_item.item20 = (My_decimal(get_last_info(tmp_item.item2, 20)) + My_decimal(5000))
        else:
            tmp_item.item20 = My_decimal(5000)
        # 累计专项扣除21
        if find_last_info(tmp_item.item2):
            tmp_item.item21 = (My_decimal(get_last_info(tmp_item.item2, 21)) + 
                                 #My_decimal(tmp_item.item10) + 
                                 My_decimal(tmp_item.item11) + 
                                 My_decimal(tmp_item.item12) + 
                                 My_decimal(tmp_item.item13) + 
                                 My_decimal(tmp_item.item14) + 
                                 My_decimal(tmp_item.item15) + 
                                 My_decimal(tmp_item.item16) + 
                                 My_decimal(tmp_item.item17) + 
                                 My_decimal(tmp_item.item18))
        else:
            tmp_item.item21 = (#My_decimal(tmp_item.item10) + 
                                 My_decimal(tmp_item.item11) + 
                                 My_decimal(tmp_item.item12) + 
                                 My_decimal(tmp_item.item13) + 
                                 My_decimal(tmp_item.item14) + 
                                 My_decimal(tmp_item.item15) + 
                                 My_decimal(tmp_item.item16) + 
                                 My_decimal(tmp_item.item17) + 
                                 My_decimal(tmp_item.item18))
        # 累计子女教育支出扣除22
        if find_last_info(tmp_item.item2):
            tmp = get_deduction_value(tmp_item.item2, "子女教育支出扣除")
            tmp_item.item22 = (My_decimal(get_last_info(tmp_item.item2, 22)) + My_decimal(tmp))
        else:
            tmp_item.item22 = My_decimal(get_deduction_value(tmp_item.item2, "子女教育支出扣除"))
        # 累计赡养老人支出扣除23
        if find_last_info(tmp_item.item2):
            tmp = get_deduction_value(tmp_item.item2, "赡养老人支出扣除")
            tmp_item.item23 = (My_decimal(get_last_info(tmp_item.item2, 23)) + My_decimal(tmp))
        else:
            tmp_item.item23 = My_decimal(get_deduction_value(tmp_item.item2, "赡养老人支出扣除"))
        # 累计继续教育支出扣除24
        if find_last_info(tmp_item.item2):
            tmp = get_deduction_value(tmp_item.item2, "继续教育支出扣除")
            tmp_item.item24 = (My_decimal(get_last_info(tmp_item.item2, 24)) + My_decimal(tmp))
        else:
            tmp_item.item24 = My_decimal(get_deduction_value(tmp_item.item2, "继续教育支出扣除"))
        # 累计住房贷款利息支出扣除25
        if find_last_info(tmp_item.item2):
            tmp = get_deduction_value(tmp_item.item2, "住房贷款利息支出扣除")
            tmp_item.item25 = (My_decimal(get_last_info(tmp_item.item2, 25)) + My_decimal(tmp))
        else:
            tmp_item.item25 = My_decimal(get_deduction_value(tmp_item.item2, "住房贷款利息支出扣除"))
        # 累计住房租金支出扣除26
        if find_last_info(tmp_item.item2):
            tmp = get_deduction_value(tmp_item.item2, "住房租金支出扣除")
            tmp_item.item26 = (My_decimal(get_last_info(tmp_item.item2, 26)) + My_decimal(tmp))
        else:
            tmp_item.item26 = My_decimal(get_deduction_value(tmp_item.item2, "住房租金支出扣除"))   
        # 累计其他扣除27
        if find_last_info(tmp_item.item2):
            tmp_item.item27 = My_decimal(0) 
        else:
            tmp_item.item27 = My_decimal(0) 
        # 累计准予扣除的捐赠28
        if find_last_info(tmp_item.item2):
            tmp_item.item28 = (My_decimal(get_last_info(tmp_item.item2, 28)) + My_decimal(0))
        else:
            tmp_item.item28 = My_decimal(0)
        # 累计应纳税所得额29
        tmp_item.item29 = (My_decimal(tmp_item.item19) - 
                           My_decimal(tmp_item.item20) - 
                           My_decimal(tmp_item.item21) - 
                           My_decimal(tmp_item.item22) - 
                           My_decimal(tmp_item.item23) - 
                           My_decimal(tmp_item.item24) - 
                           My_decimal(tmp_item.item25) - 
                           My_decimal(tmp_item.item26) -
                           My_decimal(tmp_item.item27) -
                           My_decimal(tmp_item.item28)
                           )

        if (tmp_item.item29 < 0):
            tmp_item.item29 = My_decimal(0)
            tmp_item.item37 = '-'
        # 税率30
        if (tmp_item.item29 <= 36000):
            tmp_item.item30 = My_decimal(3)
        elif (tmp_item.item29 <= 144000):
            tmp_item.item30 = My_decimal(10)
        elif (tmp_item.item29 <= 300000):
            tmp_item.item30 = My_decimal(20)  
        elif (tmp_item.item29 <= 420000):
            tmp_item.item30 = My_decimal(25) 
        elif (tmp_item.item29 <= 660000):
            tmp_item.item30 = My_decimal(30)
        elif (tmp_item.item29 <= 960000):
            tmp_item.item30 = My_decimal(35) 
        else:
            tmp_item.item30 = My_decimal(45)  
        # 速算扣除数31
        tmp_item.item31 = My_decimal(0)
        if (tmp_item.item29 <= 36000):
            tmp_item.item31 = My_decimal(0)
        elif (tmp_item.item29 <= 144000):
            tmp_item.item31 = My_decimal(2520)
        elif (tmp_item.item29 <= 300000):
            tmp_item.item31 = My_decimal(16920)
        elif (tmp_item.item29 <= 420000):
            tmp_item.item31 = My_decimal(31920) 
        elif (tmp_item.item29 <= 660000):
            tmp_item.item31 = My_decimal(52920)
        elif (tmp_item.item29 <= 960000):
            tmp_item.item31 = My_decimal(85920) 
        else:
            tmp_item.item31 = My_decimal(181920)  
        # 累计应纳税额  32        
        tmp_item.item32 = My_decimal((My_decimal(tmp_item.item29) * (My_decimal(tmp_item.item30)/100) - 
                                My_decimal(tmp_item.item31)))
        # 累计减免税额33
        tmp_item.item33 = (My_decimal(get_last_info(tmp_item.item2, 33)) + My_decimal(0))
        # 累计应扣缴税额34
        tmp_item.item34 = (My_decimal(tmp_item.item32) - My_decimal(tmp_item.item33))
        # 累计已预缴税额35
        tmp_item.item35 = (My_decimal(get_last_info(tmp_item.item2, 35)) + 
                                My_decimal(get_last_info(tmp_item.item2, 36)))
        # 累计应补(退)税额36
        tmp_item.item36 = (My_decimal(tmp_item.item34) - My_decimal(tmp_item.item35))
        if tmp_item.item36 < 0:
            tmp_item.item36 = My_decimal(0)
        # 备注37
        if find_last_info(tmp_item.item2) == 0:
            tmp_item.item37 = '新增'
        
        g_person_list.append(tmp_item)
        
def generate_file():
    global g_person_list
    global g_input_this_file_dir
    global g_output_file_path
    global g_input_year_month
    
    if (check_warning_year_month()== 0):
        return
    
    if (check_warning_input_last_file()== 0):
        return
    
    if (check_warning_input_this_file()== 0):
        return
    
    merge_person_info()
    print("start generate excel fiel")
    g_output_file_path = g_input_this_file_dir + str(g_input_year_month) + '_税款计算.xls'
    work_book = xlwt.Workbook()
    sheet_type1 = work_book.add_sheet(str(g_input_year_month))
    
    borders = xlwt.Borders()  # Create borders
    borders.left = xlwt.Borders.THIN  # 添加边框-虚线边框
    borders.right = xlwt.Borders.THIN  # 添加边框-虚线边框
    borders.top = xlwt.Borders.THIN  # 添加边框-虚线边框
    borders.bottom = xlwt.Borders.THIN  # 添加边框-虚线边框
    borders.left_colour = 0x00 # 边框上色
    borders.right_colour = 0x00
    borders.top_colour = 0x00
    borders.bottom_colour = 0x00

    style_border = xlwt.XFStyle()  # Create style
    style_border.borders = borders  # Add borders to style

    # add new colour to palette and set RGB colour value
    xlwt.add_palette_colour("custom_colour", 0x21)
    work_book.set_colour_RGB(0x21, 146, 208, 80)
    style_pattern_green = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')
    style_pattern_green.borders = borders
    
    # add new colour to palette and set RGB colour value
    xlwt.add_palette_colour("custom_colour", 0x22)
    work_book.set_colour_RGB(0x22, 252, 228, 214)
    style_pattern_orange = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')
    style_pattern_orange.borders = borders
    
    for col in range(len(title_gen)):
        sheet_type1.write(0, col, title_gen[col], style_border)
    
    for index in range(len(g_person_list)):
        tmp_col = 0
        style = style_border
        if(g_person_list[index].item37 == "-"):
            style = style_border
            g_person_list[index].item37 = ""
        elif (g_person_list[index].item37 == ""):
            style = style_pattern_orange
        else:
            style = style_pattern_green
        for str_item in vars(g_person_list[index]).keys():
            sheet_type1.write(index+1, tmp_col, vars(g_person_list[index])[str_item], style)
            tmp_col = tmp_col + 1
    
    work_book.save(g_output_file_path)
    var_show_root_output_file.set('生成文件： ' + g_output_file_path)
    print('end write excel file')

def start_app():
    global btn_input_last_file
    global btn_input_this_file
    global btn_input_deduction_file
    
    root_main.title('纳税生成')
    root_main.geometry('700x655')
    root_str.place(x=10,y=10)
    lbl_pay_total.place(x=30,y=10)
    txt_pay_total.place(x=200,y=10)
    
    lbl_yanglao.place(x=30,y=40)
    txt_yanglao.place(x=200,y=40)
    
    lbl_yiliao.place(x=30,y=70)
    txt_yiliao.place(x=200,y=70)
    
    lbl_shiye.place(x=350,y=10)
    txt_shiye.place(x=520,y=10)
    
    lbl_gongjijin.place(x=350,y=40)
    txt_gongjijin.place(x=520,y=40)
    
    lbl_this_start.place(x=30,y=100)
    txt_this_start.place(x=200,y=100)
    
    lbl_this_end.place(x=250,y=100)
    txt_this_end.place(x=280,y=100)
    
    root_input.place(x=10,y=190)
    lbl_input_year_month.place(x=30, y=10)
    txt_input_year_month.place(x=350, y=10)

    btn_input_deduction_file = Button(root_input, text="选择专项扣除文件", width=75, command=select_input_deduction_file)
    btn_input_deduction_file.place(x=30, y=45)
    lbl_show_input_deducation_file.place(x=30, y=85)
    
    btn_input_last_file = Button(root_input, text="选择上月税款计算文件", width=75, command=select_input_last_file)
    btn_input_last_file.place(x=30, y=120)
    lbl_show_input_last_file.place(x=30, y=165)
    
    btn_input_this_file = Button(root_input, text="选择本月工资明细文件", width=75, command=select_input_this_file)
    btn_input_this_file.place(x=30, y=195)
    lbl_show_input_this_file.place(x=30, y=240)
    
    root_output.place(x=10,y=505)
    btn_output = Button(root_output, text="生成本月纳税文件", width=75, command=generate_file)
    btn_output.place(x=30, y=10)
    lbl_show_root_output_file.place(x=30, y=55)
    
    root_main.mainloop()

def get_input_str():
    global g_str_this_total
    global g_str_this_yanglao
    global g_str_this_shiye
    global g_str_this_gongjijin
    global g_str_this_fangbu
    global g_input_this_file_start_index
    global g_input_this_file_end_index
    global g_str_this_tax_arrears
    
    g_str_this_total = txt_pay_total.get()
    g_str_this_yanglao = txt_yanglao.get()
    g_str_this_shiye = txt_shiye.get()
    g_str_this_gongjijin = txt_gongjijin.get()
    g_input_this_file_start_index = int(txt_this_start.get())
    g_input_this_file_end_index = int(txt_this_end.get())

if __name__ == '__main__':
    start_app()
