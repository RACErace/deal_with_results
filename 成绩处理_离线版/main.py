import xlrd
import xlwt
import sys, os
import filetype
import Ui_untitled
from shutil import *
from xlutils.copy import copy
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow, QPushButton

def dlrst():
    kind1 = filetype.guess(ui.label.tittle)
    kind2 = filetype.guess(ui.label_2.tittle)

    if (kind1.extension or kind2.extension) != 'xls':
        print ('error')
    else:
        #寻找目标所在行数
        def searchrow(target,sheetname):
            a=sheetname.nrows
            for b in range(a):  
                for c in sheetname.row_values(b):
                    if c == target:
                        return b

        #寻找目标所在列数
        def searchcol(target,sheetname):
            a=sheetname.ncols
            for b in range(a):  
                for c in sheetname.col_values(b):
                    if c == target:
                        return b

        #以字典的形式输出存在的科目及其所在位置
        def printsubjectdict(sheetname):
            possiblesubject = {"总分","语文","数学","理数","数学(理)","数学（理）","理科数学","英语","物理","化学","生物","理综","理科综合"}
            subjectdic = {}
            for c in range(3):
                searchlist = sheetname.row_values(c)
                reallysubject = possiblesubject&set(searchlist)
                if len(reallysubject)>0:
                    for b in list(reallysubject):
                        subjectdic[b] = searchcol(b,sheetname)
                    return subjectdic
            return 0

        #以列表的形式输出存在的科目
        def printsubjectlist(sheetname):
            possiblesubject = {"总分","语文","数学","理数","数学(理)","数学（理）","理科数学","英语","物理","化学","生物","理综","理科综合"}
            a=sheetname.nrows
            for b in range(a):
                searchlist = sheetname.row_values(b)
                reallysubject = possiblesubject&set(searchlist)
                if len(reallysubject) > 0:
                    return list(reallysubject) 

        #以字典的形式输名字所在的位置
        def printnamedict(oldfilename):
            oldnamelist = oldfilename.sheet_names()
            namedict = {}
            for a in range(len(oldnamelist)):
                namedict[oldnamelist[a]] = a
            return namedict

        #以列表的形式输出可用的名单
        def printnamelist(sourcesheetname,oldfilename):
            oldnamelist = oldfilename.sheet_names()
            a=sourcesheetname.ncols
            for b in range(a):
                for c in sourcesheetname.col_values(b):
                    if c in oldnamelist:
                        sourcesnamelist = sourcesheetname.col_values(b)
                        namelist = set(oldnamelist)&set(sourcesnamelist)
                        return list(namelist)

        standardsubjectdict = {"总分":2,"语文":5,"数学":8,"理数":8,"数学(理)":8,"数学（理）":8,"理科数学":8,"英语":11,"物理":14,"化学":17,"生物":20,"理综":23,"理科综合":23}
        if os.path.exists(os.path.dirname(os.path.realpath(__file__)) + r'\backup'):
            copy2(ui.label.tittle, os.path.dirname(os.path.realpath(__file__)) + r'\backup\成绩对比表.xls')
        else:
            os.makedirs(os.path.dirname(os.path.realpath(__file__)) + r'\backup')
            open(os.path.dirname(os.path.realpath(__file__)) + r'\backup\成绩对比表.xls', mode='wb+')
            copy2(ui.label.tittle, os.path.dirname(os.path.realpath(__file__)) + r'\backup\成绩对比表.xls')
        sourcefile = xlrd.open_workbook(ui.label_2.tittle) #打开成绩表文件
        sourcesheet = sourcefile.sheets()[0] #获取成绩表文件中第一个sheet
        oldfile = xlrd.open_workbook(ui.label.tittle) #打开已有的成绩对比表文件
        examname = sourcefile.sheet_names() #从成绩表文件中第一个sheet获取考试名称
        newfile = copy(oldfile) #将xlrd对象拷贝转化为xlwt对象
        
        subjectdict = printsubjectdict(sourcesheet) #以字典的形式输出存在的科目及其所在位置
        subjectlist = printsubjectlist(sourcesheet) #以列表的形式输出存在的科目
        namedict = printnamedict(oldfile) #以字典的形式输名字所在的位置
        namelist = printnamelist(sourcesheet,oldfile) #以列表的形式输出可用的名单

        #计算成绩表文件中各项的间距
        position = list(subjectdict.values())
        position.sort()
        space = position[1] - position[0]
        if space > 3:
            space = 3
        else:
            space = space
        
        for name in namelist:
            oldsheet = oldfile.sheet_by_name(name) #通过sheet名称获取需要操作的sheet
            row = oldsheet.nrows #获取sheet中有效行数
            newsheet = newfile.get_sheet(namedict[name]) #获取需要操作的sheet
            newsheet.write(row, 0, examname) # 写入考试名称
            newsheet.write(row, 1, name) # 写入名称
            rowname = searchrow(name,sourcesheet) #获取名字所在的行
            for subject in subjectlist:
                for a in range(space):
                    value = sourcesheet.cell_value(rowx = rowname, colx = subjectdict[subject] + a)
                    newsheet.write(row, standardsubjectdict[subject]+a, value) # 写入excel, 参数对应 行, 列, 值

        newfile.save('成绩对比表.xls')  # 保存工作簿

def create():
    kind = filetype.guess(ui.label_3.tittle)
    if kind.extension != 'xls':
        print ('error!')
    else:
        namebook = xlrd.open_workbook(ui.label_3.tittle) #打开工作薄
        namesheet = namebook.sheets()[0] #获取第一个sheet表格
        namelist = namesheet.col_values(colx=0, start_rowx=0, end_rowx=None)
        namelist = list(set(namelist))
        namelistfile = xlwt.Workbook(encoding = 'utf-8') #创建一个workbook并设置编码
        for name in namelist:
            sheet = namelistfile.add_sheet(name) #添加sheet
            sheet.write(0, 0, '考试名称') # 写入excel, 参数对应 行, 列, 值
            sheet.write(0, 1, '姓名')
            sheet.write(0, 2, '总分')
            sheet.write(0, 5, '语文')
            sheet.write(0, 8, '数学')
            sheet.write(0, 11, '英语')
            sheet.write(0, 14, '物理')
            sheet.write(0, 17, '化学')
            sheet.write(0, 20, '生物')
            sheet.write(0, 23, '理综')
            sheet.write(1, 2, '得分')
            sheet.write(1, 3, '校次')
            sheet.write(1, 4, '班次')
            sheet.write(1, 5, '得分')
            sheet.write(1, 6, '校次')
            sheet.write(1, 7, '班次')
            sheet.write(1, 8, '得分')
            sheet.write(1, 9, '校次')
            sheet.write(1, 10, '班次')
            sheet.write(1, 11, '得分')
            sheet.write(1, 12, '校次')
            sheet.write(1, 13, '班次')
            sheet.write(1, 14, '得分')
            sheet.write(1, 15, '校次')
            sheet.write(1, 16, '班次')
            sheet.write(1, 17, '得分')
            sheet.write(1, 18, '校次')
            sheet.write(1, 19, '班次')
            sheet.write(1, 20, '得分')
            sheet.write(1, 21, '校次')
            sheet.write(1, 22, '班次')
            sheet.write(1, 23, '得分')
            sheet.write(1, 24, '校次')
            sheet.write(1, 25, '班次')
        namelistfile.save('成绩对比表.xls') #保存

if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_untitled.Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    ui.pushButton.clicked.connect(dlrst)
    ui.pushButton_2.clicked.connect(create)
    sys.exit(app.exec_())