#dlrst_1.1  by:race

import xlrd
import xlwt
import os
from xlutils.copy import copy
from pywebio.input import *
from pywebio.output import *
from pywebio import session, start_server
import filetype
import time

def index ():
    put_row([put_buttons(['处理成绩'], [lambda: session.go_app('dlrst')]), put_buttons(['生成成绩对比表模板'], [lambda: session.go_app('create')])])
    put_text('使用说明：')
    put_text('1.此应用只支持“.xls”格式的文件。')
    put_text('2.若是第一次使用此应用，请先点击“生成成绩对比表模板”，通过上传名单来获得绩对比表模板。')
    var = 1
    while var == 1:
        nowtime = time.time()
        massage = input(label='问题反馈')
        a = '[问题反馈]' + nowtime + ' ' + massage
        print (a)
        with open("matter.txt","a") as file:
            file.write(a + '\n')

def dlrst():
    put_text('个人成绩对比表生成器 by:Race').style('color:#C0C0C0 ; font-size: 5px; text-align: center')
    put_text('说明：请先上传成绩表文件和已有的成绩对比表文件，点击“下载”后，再点击“完成”，以便再次使用本应用。').style('color:##1473FF ; font-size: 15px')
    put_text('注意：请将成绩表文件中工作表的名称更改为考试名称').style('color:#FF0000 ; font-size: 15px')

    #上传成绩表文件
    sourcebook = file_upload("上传成绩表", placeholder="请选择文件", accept=".xls", required=True)
    nowtime1 = str(time.time())
    open(nowtime1 + sourcebook["filename"],"wb").write(sourcebook['content'])

    #上传成绩对比表文件
    oldbook = file_upload("上传成绩对比表", placeholder="请选择文件", accept=".xls", required=True)
    nowtime2 = str(time.time())
    open(nowtime2 + oldbook["filename"],"wb").write(oldbook['content'])

    kind1 = filetype.guess(nowtime1 + sourcebook["filename"])
    kind2 = filetype.guess(nowtime2 + oldbook["filename"])

    if (kind1.extension or kind2.extension) != 'xls':
        os.remove(nowtime1 + sourcebook["filename"])
        os.remove(nowtime2 + oldbook["filename"])
        put_text('格式错误！').style('color:#FF0000')
        put_text('请不要通过直接更改文件后缀名来更改文件格式，直接更改文件后缀名将导致文件无法识别！').style('color:#FF0000')
        put_text('请通过将文件“另存为”来更改文件格式！').style('color:#FF0000')
        put_button("确定", onclick=lambda: session.run_js('location.reload()'))
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
        sourcefile = xlrd.open_workbook(nowtime1 + sourcebook["filename"]) #打开成绩表文件
        sourcesheet = sourcefile.sheets()[0] #获取成绩表文件中第一个sheet
        oldfile = xlrd.open_workbook(nowtime2 + oldbook["filename"]) #打开已有的成绩对比表文件
        examname = sourcefile.sheet_names() #从成绩表文件中第一个sheet获取考试名称
        newfile = copy(oldfile) #将xlrd对象拷贝转化为xlwt对象
        
        subjectdict = printsubjectdict(sourcesheet) #以字典的形式输出存在的科目及其所在位置
        subjectlist = printsubjectlist(sourcesheet) #以列表的形式输出存在的科目
        namedict = printnamedict(oldfile) #以字典的形式输名字所在的位置
        namelist = printnamelist(sourcesheet,oldfile) #以列表的形式输出可用的名单
        print (subjectdict)
        print (subjectlist)
        print (namedict)
        print (namelist)

        #计算成绩表文件中各项的间距
        position = list(subjectdict.values())
        print (position)
        position.sort()
        print (position)
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

        nowtime3 = str(time.time())
        newfile.save(nowtime3 + oldbook["filename"])  # 保存工作簿

        content = open(nowtime3 + oldbook["filename"], 'rb').read()
        put_file(oldbook["filename"], content, '下载')
        def finish():
            os.remove(nowtime1 + sourcebook["filename"])
            os.remove(nowtime2 + oldbook["filename"])
            os.remove(nowtime3 + oldbook["filename"])
            session.run_js('location.reload()')
        put_button("完成", onclick=finish)

def create():
    put_text('个人成绩对比表生成器 by:Race').style('color:#C0C0C0 ; font-size: 5px; text-align: center')
    put_text('提示：处理结束后先点击“下载”再点击“去处理成绩”。')
    namesourcefile = file_upload("上传名单文件", accept=".xls",placeholder="选择一个xls表格文件",multiple=False,required=True)
    nowtime = str(time.time())
    open(nowtime + namesourcefile['filename'], 'wb').write(namesourcefile['content'])
    kind = filetype.guess(nowtime + namesourcefile['filename'])
    if kind.extension != 'xls':
        os.remove(nowtime + namesourcefile['filename'])
        put_text('格式错误！').style('color:#FF0000')
        put_text('请不要通过直接更改文件后缀名来更改文件格式，直接更改文件后缀名将导致文件无法识别！').style('color:#FF0000')
        put_text('请通过将文件“另存为”来更改文件格式！').style('color:#FF0000')
        put_button("确定", onclick=lambda: session.run_js('location.reload()'))
    else:
        namebook = xlrd.open_workbook(nowtime + namesourcefile['filename']) #打开工作薄
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
        namelistfile.save(nowtime + 'namelistfile.xls') #保存
        content = open(nowtime + 'namelistfile.xls', 'rb').read()
        put_file('成绩对比表.xls', content, '下载')
        def finish():
            os.remove(nowtime + namesourcefile['filename'])
            os.remove(nowtime + 'namelistfile.xls')
            session.run_js('location.reload()')        
        put_buttons(['去处理成绩'], [lambda: session.go_app('dlrst', new_window=False)])
        session.defer_call(finish)
    
if __name__ == '__main__':
    start_server([index, create, dlrst], port=888)
