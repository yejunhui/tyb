import xlrd,xlwt
import re
import os
import time

class ExcelPlan(object):
	"""docstring for ExcelRead"""
	def __init__(self, arg):
		super(ExcelPlan, self).__init__()
		self.arg = arg
		self.titleExcel = '物料清单模板'
		self.tips = '''本模板用于导入报表数据\n
					   请不要更改模板格式\n
					   你可以拷贝本模板到任何.xls文件中，但请记住本模板表名必需为Template（区分大小写）且你的文件中没有同名的工作表\n
					   数据请按以下规则录入
					'''
		self.time = time.ctime()
		self.title = ['类','ERP编码','名称','模号','用量','单位','备注','盘点']
		self.plan = ['计划生产','计划出货','库存']
	
	"""写入excel模板"""
	def writeExcel(self):
		f = xlwt.Workbook()
		sheet1 = f.add_sheet('Template',cell_overwrite_ok=True)
		style1 = xlwt.XFStyle()
		row = 0
		col = 0
		sheet1.write(row,col,self.titleExcel)
		row +=1
		sheet1.write(row,col,self.tips)
		row +=2
		sheet1.write(row,col,self.title)
		for t in self.title :
			sheet1.write(row,col,t)
			col +=1
		if not os.path.exists(self.arg+'/Template.xls') :
			f.save(self.arg+'/Template.xls')
		else :
			os.remove(self.arg+'/Template.xls')
			f.save(self.arg+'/Template.xls')


	"""写入计划excel"""
	def writePlan(self,data):
		f = xlwt.Workbook()
		sheet1 = f.add_sheet('Plan',cell_overwrite_ok=True)
		row = 0
		col = 0
		"""写表头"""
		sheet1.write(row,col,"组装物料")
		row +=1
		sheet1.write(row,col,"更新时间%s" %self.time)
		row +=1
		for t in self.title :
			sheet1.write(row,col,t)
			col +=1
		for i in range(31) :
			sheet1.write(row,col,i+1)
			col +=1
		
		row +=1
		col = 0
		# 定义一号所在的列
		dcol = 8
		# 定义用量所在列
		ccol = 4
		# 写数据
		for d in range(data) :
			if d[0] =='总成' :
				frow = row
				for p in range(self.plan) :
					for o in range(d) :
						sheet1.write(row,col,o)
						col += 1
					sheet1.write(row,col,p)
					row +=1
				col = 0
			else :
				for p in range(self.plan) :
					for o in range(d) :
						sheet1.write(row,col,o)
						col += 1
					sheet1.write(row,col,p)
					col += 1
				for i in range(31) :
					fc = '=ADDRESS(frow,)'
					sheet1.write_formula(row,col,'')
					col += 1
				col = 0

		f.save(self.arg+'/PlanExcel.xls')


	"""读取excel文件"""
	def readExcel(self,file):
		book = xlrd.open_workbook(file)
		sheet1 = book.sheet_by_name('Template')
		rows = sheet1.nrows
		row = 0
		col = 0
		data = []
		old = []
		j = 0
		for i in range(rows) :
			li = sheet1.row_values(i)
			if li[0] == '总成' :
				if j != 0 :
					print('将%s写入data'%li[1])
					data.append(old)
					old = []
				print('diyiczc')
				j +=1
			if li[1] != '' and li[1] != 'ERP编码' :
				old.append(li)
			
		return data

