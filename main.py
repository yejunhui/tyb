from xlsx import ExcelPlan
import os
import time

#录入目录
p1 = input('请输入目录(直接回车使用当前目录)：')
if p1 == '':
    p1 = os.getcwd()
else:
    pass
print('当前工作目录：%s'%p1)
#选择操作
print('-'*100)
print(time.localtime(time.time()).tm_year,'/',time.localtime(time.time()).tm_mon,'/',time.localtime(time.time()).tm_mday,'',time.localtime(time.time()).tm_hour,':',time.localtime(time.time()).tm_min,'/',time.localtime(time.time()).tm_sec)
print('-'*100)
print('1、生成清单模板\n')
print('2、生产推移表\n')
print('0、退出\n')
print('-'*100)
option = input('请选择：')

e = ExcelPlan(p1)

if option == '0':
    exit()
elif option == '1':
    e.writeExcel()
elif option == '2':
    p2 = input('请输入文件名：')

    d,n = e.readExcel(p1+'\\'+p2)

    #传入列表，在定义的目录生产跟进表"PlanExcel.xlsx"
    e.writePlan(d,n)
else:
    print('输入错误')
    exit()
