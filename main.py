import shutil

import xlrd  #读表格
import xlwt  #写表格
from xlutils.copy import copy #表格备份
from tkinter import * #GUI
from tkinter import filedialog
import tkinter.messagebox #弹出提示
import time #输出时间
import os #用于检测文件是否存在
import platform
import subprocess
import glob
# input_list=[]
# xlsrowvalue=[]
global xlsrowvalue
global input_list
input_list=[]
xlsrowvalue=[]

global save_path
save_path=''

global open_file
open_file=''



def GetDesktopPath():
    return os.path.join(os.path.expanduser("~"), 'Desktop')


def askPath():
    global save_path
    root = Tk()
    root.withdraw()
    save_path = filedialog.askdirectory()
    # return:
    # f_path

def askFile():
    global open_file
    root = Tk()
    root.withdraw()
    open_file = filedialog.askopenfilename()

def mycopy():
    askFile()
    global open_file
    # paths=open_file
    shutil.copy(open_file, './'+'名单.xls')


def inputList1():
    global xlsrowvalue
    global input_list
    root = Tk()
    root.withdraw()
    f_path = filedialog.askopenfilename()
    xls_file = xlrd.open_workbook(f_path)
    temp = choose_sheet.get()
    xls_sheet = xls_file.sheets()[0]
    xlsrowvalue = []
    for i in range(xls_sheet.nrows):
        row_value = xls_sheet.row_values(i)
        xlsrowvalue.append(str(row_value[0]))
    input_list =xlsrowvalue
    # print(input_list)
    tkinter.messagebox.showinfo('提示', '签到详情已导入')
def inputList2():
    global input_list
    global xlsrowvalue
    a = text_box.get()
    # print(a.split(' '))
    a = a.split(' ')
    for i in a:
        input_list.append(i)
    # input_list.append(a.split(' '))
    tkinter.messagebox.showinfo('提示', '签到详情已写入')


def type():
    choose = var.get()
    # tkinter.messagebox.showinfo('提示', '切换成功')

def clean():
    text_box.delete(0, END)


# 文本输出
def main():
    global name_list
    global xlsrowvalue
    global input_list
    out_box.delete(1.0, END)
    if os.path.exists("名单.xls") == False:
        a = tkinter.messagebox.askokcancel('提示', '您还没有名单''\n''需要创建吗')
        if a == True:

            create = xlwt.Workbook()
            worksheet = create.add_sheet('姓名')
            worksheet.write(0, 0, '姓名')
            worksheet.write(1, 0, ' ')
            create.save('名单.xls')
            tkinter.messagebox.showinfo('提示', '名单已创建''\n''请前往编辑')
        else:
            tkinter.messagebox.showinfo('提示', '操作已取消')
    else:
        read_book = xlrd.open_workbook('名单.xls', formatting_info=True)
        temp =choose_sheet.get()
        if temp=='4':
            main_data = read_book.sheets()[1]
        else:
            main_data = read_book.sheets()[0]
        name_list = main_data.col_values(0)
        # print(name_list)
        if main_data.cell(1, 0).value == ' ':
            tkinter.messagebox.showinfo('提示', '您还未添加学生信息')
        else:
            # 取得输入数据
            # input_list = xlsrowvalue

            ontime_list = []
            late_list = []

            member_ontime = 0
            member_late = 0

            i = 0
            choose = var.get()
            if choose == 0:
                while i < len(name_list):
                    if name_list[i] in input_list:
                        txt = name_list[i] + '  √''\n\n'
                        out_box.insert(END, txt)
                    else:
                        txt = name_list[i] + '  X' + '\n\n'
                        out_box.insert(END, txt)
                    i = i + 1

            elif choose == 1:
                while i < len(name_list):
                    if name_list[i] in input_list:
                        ontime_list.append(1)
                        ontime_list[member_ontime] = name_list[i]
                        member_ontime = member_ontime + 1
                    else:
                        late_list.append(1)
                        late_list[member_late] = name_list[i]
                        member_late = member_late + 1
                    i = i + 1
                out_box.insert(END, '准时签到名单')
                out_box.insert(END, '\n')
                ontime = ' '.join(ontime_list) + '\n' + '\n'
                out_box.insert(END, ontime)
                out_box.insert(END, '\n')
                out_box.insert(END, '迟到名单')
                out_box.insert(END, '\n')
                late = ' '.join(late_list) + '\n'
                out_box.insert(END, late)


# 表格保存
def save():
    global name_list
    global input_list
    global save_path
    if os.path.exists("名单.xls") == False:
        a = tkinter.messagebox.askokcancel('提示', '您还没有创建名单''\n''需要创建吗')
        if a == True:
            # 定义一个输入文本框
            # entry = tk.Entry(window, show="*")
            # 表示输入的字符以*号的形式出现

            worksheet = create.add_sheet('名单')
            worksheet.write(0, 0, '姓名')
            worksheet.write(1, 0, ' ')
            create.save('名单.xls')
            tkinter.messagebox.showinfo('提示', '名单已创建''\n''请前往编辑')
        else:
            tkinter.messagebox.showinfo('提示', '操作已取消')
    else:
        read_book = xlrd.open_workbook('名单.xls', formatting_info=True)
        main_data = read_book.sheets()[0]

        if main_data.cell(1, 0).value == ' ':
            tkinter.messagebox.showinfo('提示', '您还未添加学生信息')
        else:
            # 读取表格
            # name_list = main_data.col_values(0)
            write_place = main_data.ncols
            write_high = main_data.nrows

            # 取得输入数据
            # input_list = xlsrowvalue
            state_list = []  # 每人签到# 状态
            i = 0
            count = 0
            # absentr = xlrd.open_workbook('')
            absent = xlwt.Workbook(encoding='utf-8')
            absent_s1 = absent.add_sheet('sheet1', cell_overwrite_ok=True)
            # for i in range()
            while i < len(name_list):
                if name_list[i] in input_list:
                    state_list.append(1)
                    state_list[i] = '√'
                else:
                    state_list.append(1)
                    state_list[i] = 'X'
                    # print(name_list[i])
                    absent_s1.write(count,0, name_list[i])
                    count += 1
                i = i + 1
            systemType: str = platform.platform()

            if address.get():
                address1=address.get()
                path=address1+'/absent.xls'
                # path=open_fp(path)

                if 'mac' in systemType:

                    path: str = path.replace("\\", "/")  # mac系统下,遇到`\\`让路径打不开,不清楚为什么哈,觉得没必要的话自己可以删掉啦,18行那条也是
                    absent.save(path)
                else:
                    path: str = path.replace("/", "\\")
                    absent.save(path)
            elif save_path!='':
                path = save_path + '/absent.xls'
                # path=open_fp(path)
                if 'mac' in systemType:

                    path: str = path.replace("\\", "/")  # mac系统下,遇到`\\`让路径打不开,不清楚为什么哈,觉得没必要的话自己可以删掉啦,18行那条也是
                    absent.save(path)
                else:
                    path: str = path.replace("/", "\\")
                    absent.save(path)
            else:
                path=GetDesktopPath()+'/absent.xls'
                print(save_path)
                if 'mac' in systemType:  # 判断以下当前系统类型
                    path: str = path.replace("\\", "/")  # mac系统下,遇到`\\`让路径打不开,不清楚为什么哈,觉得没必要的话自己可以删掉啦,18行那条也是
                    absent.save(path)
                    subprocess.call(["open", path])
                else:
                    path: str = path.replace("/", "\\")
                    absent.save(path)
                    os.startfile(path)

            new_excel = copy(read_book)
            ws = new_excel.get_sheet(0)
            i = 1

            while i <= write_high:
                try:
                    ws.write(i - 1, write_place, state_list[i - 1])  # 写入签到状态
                    i = i + 1
                except:
                    break
            time_now = time.strftime("%m-%d %H:%M", time.localtime())
            ws.write(0, write_place, time_now)  # 写入时间
            new_excel.save('名单.xls')
        tkinter.messagebox.showinfo('提示', '已导出表格到'+save_path+'，请查看')


root = Tk()
root.geometry('460x300')
root.title('zhoukeng签到')
root.iconbitmap(r'./logo.ico')


tkinter.messagebox.showinfo('提示', '如果桌面已经有absent.xls，请删除')

text_box = Entry(root)
text_box.place(x=10, y=10, height=20, width=360)

start_btn = Button(root, text='确认', command=inputList2)
start_btn.place(x=375, y=10, height=20, width=80)

start_btn = Button(root, text='导入表格', command=inputList1)
start_btn.place(x=375, y=35, height=20, width=80)

start_btn = Button(root, text='运行', command=main)
start_btn.place(x=375, y=60, height=20, width=80)

out_box = Text(root)
out_box.place(x=10, y=40, height=160, width=360)

title_choose = Label(root,text='请输入缺席统计表保存路径（默认桌面）：')
title_choose.place(x=10, y=200)

address = Entry(root)
address.place(x=10, y=220, height=20, width=360)

address_button=Button(root, text='保存路径', command=askPath)
address_button.place(x=375,y=220,width=80,height=20)

title_choose = Label(root,text='请选择班级名单（单个数字，3班4班一起不填）：')
title_choose.place(x=10, y=240)

choose_sheet= Entry(root)
choose_sheet.place(x=10, y=260, height=20, width=360)

title_choose = Label(root,text='模式选择')
title_choose.place(x=390, y=90)

var = IntVar()
rd1 = Radiobutton(root,text="啰嗦模式",variable=var,value=0,command=type)
rd1.place(x=375, y=110)

# var = IntVar()
rd2 = Radiobutton(root,text="简洁模式",variable=var,value=1,command=type)
rd2.place(x=375, y=135)

start_btn = Button(root, text='导入名单', command=mycopy)
start_btn.place(x=375, y=160, height=20, width=80)

start_btn = Button(root, text='输出表格', command=save)
start_btn.place(x=375, y=190, height=20, width=80)

title_choose = Label(root,text='Zhoukeng')
title_choose.place(x=375, y=270,)
root.mainloop()
