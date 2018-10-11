from xlsx import ExcelPlan
#定义目录
e = ExcelPlan('f:/py/tyb')
#读取文件，提供文件名
d = e.readExcel('/Template.xlsx')
#传入列表，在定义的目录生产跟进表
e.writePlan(d)
