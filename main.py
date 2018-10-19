from xlsx import ExcelPlan

p1 = input('请输入目录：')
p2 = input('请输入文件名：')

e = ExcelPlan(p1)
d,n = e.readExcel(p1+'\\'+p2)

#传入列表，在定义的目录生产跟进表"PlanExcel.xlsx"
e.writePlan(d,n)
