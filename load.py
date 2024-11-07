# 1.读取表格中1栋数据
# 2.按2A->2B->3B->3A-...->6B排序，并附上每间个数
# 3.按顺序显示在屏幕上，加上送完按钮，以及总览界面

import openpyxl
# from tkinter import *
# import tkinter.filedialog

# xlfile = tkinter.filedialog.askopenfilename()


def loadxlfile(xlfile):
    wb = openpyxl.load_workbook(xlfile)
    sheet = wb['Sheet1']
    data = []  # 储存有效数据

    for cell_row in sheet:
        nobuilding = cell_row[0].value
        if (nobuilding == '1A' or nobuilding == '1a'):
            data.append({'A': cell_row})
        elif(nobuilding == '1B' or nobuilding == '1b'):
            data.append({'B': cell_row})

    tmp = tuple(data[0].values())
    total = tmp[0][3].value  # 应送总数
    dict = {}  # 门户字典
    for dics in data:
        for k, v in dics.items():
            section = k
            room = k+str(v[1].value)
            num = v[2].value
            if room in dict:
                dict[room] += num
            else:
                dict[room] = num

    romlist = []
    numlist = []
    with open('DORM1.txt') as f:
        sequence = f.read().splitlines()

    for item in sequence:
        if item in dict:
            romlist.append(item)
            numlist.append(dict[item])

    # lenth = len(romlist)
    # for i in range(lenth):
    #     print(romlist[i], numlist[i])

    checktotal = sum(numlist)
    print("check!")if checktotal == total else print("error!")
    return romlist, numlist
