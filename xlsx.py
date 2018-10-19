from typing import Any, Union

import xlrd
import xlsxwriter
import re
import os
import time

from xlsxwriter.worksheet import Worksheet


class ExcelPlan(object):
    """docstring for ExcelRead"""
    def __init__(self, arg):
        super(ExcelPlan, self).__init__()
        '''定义文件路径'''
        self.arg = arg
        '''定义清单模板'''
        self.titleExcel = '物料清单模板'
        self.tips = '''1.本模板用于导入报表数据
                       2.请不要更改模板格式
                       3.你可以拷贝本模板到任何.xls文件中，但请记住本模板表名必需为Template（区分大小写）且你的文件中没有同名的工作表
                       4.数据请按以下规则录入
                    '''
        '''获得系统时间'''
        self.time = time.ctime()
        '''定义计划头'''
        self.title = ['类','ERP编码','名称','模号','用量','单位','备注','盘点','']
        '''定义项目'''
        self.plan = ['计划生产','计划出货','库存']
        '''定义列标，从第9列开始'''
        self.str = ['i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','aa','ab','ac','ad','ae','af','ag','ah','ai','aj','ak','al','am','an','ao']

    """读取excel文件"""

    def readExcel(self, file):
        print('开始读取文件...')
        book = xlrd.open_workbook(file)
        sheet1 = book.sheet_by_name('Template')
        rows = sheet1.nrows
        print(rows)
        row = 0
        col = 0
        data = []
        od =[]
        newdata = []
        """读取列表"""

        i = p = 0
        for i in range(rows) :
            da = sheet1.row_values(i)
            if da[0] == '总成' :
                fn = da[1]
                if p == 1 :
                    data.append(od)
                    od = []
                    print('正在获取%s...\n' % fn)
                else:
                    p = 1
            if da[1] !='' and da[0] !='类' :
                #print('%s写入%s汇总\n'%(da,fn))
                od.append(da)
        """获取零件列表"""

        for d in data :
            for l in d :
                if l[0] != '总成' :
                    if l not in newdata :
                        newdata.append(l)
        return data,newdata

    """写入excel模板"""

    def writeExcel(self):
        f = xlsxwriter.Workbook('Template.xlsx')
        sheet1 = f.add_worksheet('Template')
        row = 0
        col = 0
        sheet1.write(row,col,self.titleExcel)
        row +=1
        sheet1.write(row,col,self.tips)
        row +=2
        sheet1.write(row,col,self.time)
        i = 0
        for t in self.title :
            if col > 7 :
                sheet1.write_formula(row,col,t)
            else :
                sheet1.write(row,col,t)
            col += 1
        f.close()

    """写入计划excel"""

    def writePlan(self,data,newdata):
        print('开始写入Plan')
        f = xlsxwriter.Workbook('PlanExcel.xlsx')
        sheet1 = f.add_worksheet('Plan')
        sheet2 = f.add_worksheet('Out')
        sheet3 = f.add_worksheet('Summary')
        row = 0
        col = 0
        sheet1.write(row,col,'组装物料跟进表')
        row += 1
        for t in self.title :
            sheet1.write(row,col,t)
            col += 1
        i = 0
        for i in range(31) :
            sheet1.write(row,col,i+1)
            col += 1
        sheet1.write(row,col,'合计')
        row += 1
        col = 0
        for d in data :
            for li in d :
                print('正在写入%s'%li[1])
                #写入总成
                if li[0] == '总成' :
                    fn ='\"'+li[1]+'\"'
                    for l in li :
                        sheet1.write(row,col,l)
                        col += 1
                    sheet1.write(row,col,self.plan[0])
                    sheet1.write_formula(row,col+32,'sum(j%s:an%s)'%(row+1,row+1))
                    sheet1.write(row+1,col,self.plan[1])
                    sheet1.write(row+2,col-1,self.plan[2])
                    col += 1
                    #写公式
                    i = 0
                    for i in range(32) :
                        sheet1.write_formula(row + 1, col + i,'vlookup(%s,Out!B:AZ,column(%s),0)' % (fn, self.str[i] + '1'))
                        sheet1.write_formula(row + 2, col + i,'indirect(address(%s,%s))+indirect(address(%s,%s))-indirect(address(%s,%s))' % ( row + 3, col+ i, row+1, col + i+1 , row + 2, col + i+1))
                    sheet1.write_formula(row+2,col+31,'indirect(address(%s,%s))'%(row+3,col+31))
                    row += 3
                    col = 0
                #写入零件
                else :
                    name = '\"'+li[1]+'\"'
                    for l in li :
                        sheet1.write(row,col,l)
                        col += 1
                    sheet1.write(row,col,self.plan[1])
                    #sheet1.write(row+1,col-1,self.plan[2])
                    col += 1
                    '''写公式'''
                    i = 0
                    for i in range(32) :
                        sheet1.write_formula(row,col+i,'vlookup(%s,B:AZ,column(%s),0)*e%s'%(fn,self.str[i]+'1',row+1))
                        #sheet1.write_formula(row+1,col+i,'vlookup(%s,Summary!B:AZ,column(%s),0)'%(name,self.str[i]+'1'))
                    row += 1
                    col = 0

        row = col = 0
        '''写入Out'''
        sheet2.write(row, col, '出货明细')
        row += 1
        sheet2.write(row,col,'本工作表只能对“计划生产”项进行操作')
        row += 1
        i = 0
        for t in self.title :
            sheet2.write(row,col,t)
            col += 1
        for i in range(31) :
            sheet2.write(row,col+i,i+1)
        sheet2.write(row,col+31,'合计')
        row += 1
        col = 0
        for d in data :
            for li in d :
                if li[0] == '总成' :
                    print('正在写入%s' % li[1])
                    for l in li :
                        sheet2.write(row,col,l)
                        col += 1
                    sheet2.write(row, col, self.plan[1])
                    sheet2.write_formula(row,col+32,'sum(j%s:an%s)'%(row+1,row+1))
                    col = 0
                    row += 1

        row = col = 0
        '''写入Summary'''
        sheet3.write(row, col, '零件明细')
        row += 1
        sheet3.write(row,col,'本工作表只能对“计划生产”项进行操作')
        row += 1
        for t in self.title :
            sheet3.write(row,col,t)
            col += 1
        i = 0
        for i in range(31) :
            sheet3.write(row,col+i,i+1)
        sheet3.write(row,col+31,'合计')
        row += 1
        col = 0
        for d in newdata :
            name = '\"'+d[1]+'\"'
            for l in d :
                 sheet3.write(row+2,col,l)
                 col += 1
            sheet3.write(row,col,self.plan[0])
            sheet3.write(row+1,col,self.plan[1])
            sheet3.write(row+2,col-1,self.plan[2])
            col += 1
            i = 0
            for i in range(32):
                sheet3.write_formula(row+1, col + i,'sumif(Plan!B:B,%s,Plan!%s:%s)'%(name,self.str[i+1],self.str[i+1]))
                sheet3.write_formula(row+2, col + i,'indirect(address(%s,%s))+indirect(address(%s,%s))-indirect(address(%s,%s))' % ( row + 3, col+ i, row+1, col + i+1, row + 2, col + i+1))
            sheet3.write_formula(row,col+31,'sum(j%s:an%s)'%(row+1,row+1))
            row += 3
            col = 0

        f.close()
        print('写入成功，文件保存在：%s'%self.arg)




