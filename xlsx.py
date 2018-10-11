import xlrd,xlsxwriter
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
		self.title = ['类','ERP编码','名称','模号','用量','单位','备注','盘点','']
		self.plan = ['计划生产','计划出货','库存']
	
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
			sheet1.write(row,col,t)
			col += 1
		f.close()



	"""写入计划excel"""
	def writePlan(self,data):
		f = xlsxwriter.Workbook('PlanExcel.xlsx')
		sheet1 = f.add_worksheet('Plan')
		sheet2 = f.add_worksheet('Out')
		row = 0
		col = 0
		print('开始写入PlanExcel.xlsx文件,文件包含Plan和Out工作表!')
		# 写Out
		for t in self.title :
			sheet2.write(row,col,t)
			col += 1
		for i in range(31) :
			sheet2.write(row,col,i+1)
			col +=1
		col = 0
		orow = row + 1
		for d in data :
			for lis in d :
				if lis[0] == '总成' :
					for li in lis :
						sheet2.write(orow,col,li)
						col += 1
					col = 0
			orow += 1
			col = 0
		"""写表头"""
		sheet1.write(row,col,"组装物料")
		row +=1
		sheet1.write(row,col,"更新时间%s" %self.time)
		row +=1
		for t in self.title :
			sheet1.write(row,col,t)
			col += 1
		for i in range(31) :
			sheet1.write(row,col,i+1)
			col +=1
		# 行加1，列重置
		row += 1
		col =0
		# 写数据
		for d in data :
			for lis in d :
				if lis[0] == '总成' :
					fname = '\"'+lis[1]+'\"'
					for li in lis :
						sheet1.write(row,col,li)
						col += 1
					ncol = col 
					sheet1.write(row,ncol,'计划生产')
					row += 1
					sheet1.write(row,ncol,'计划出货')
					i = 1
					for i in range(31) :
						i += 1
						sheet1.write_formula(row,ncol+i,'IFERROR(VLOOKUP(%s,Out!B:AB,%d+8,0),0)'%(fname,i))
					sheet1.write_formula(row,ncol+32,'SUN(INDIRECT(ADDRESS(%s,%s)):INDIRECT(ADDRESS(%s,%s)))'%(row,ncol,row,ncol+31))
					row += 1
					sheet1.write(row,ncol-1,'库存')
					i = 1
					for i in range(31) :
						i += 1
						sheet1.write_formula(row,ncol+i,'INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s))'%(row+1,ncol+i,row-1,ncol+i+1,row,ncol+i+1))
					sheet1.write_formula(row,ncol+32,'SUM(INDIRECT(ADDRESS(%s,%s)):INDIRECT(ADDRESS(%s,%s)))'%(row,ncol,row,ncol+31))
					row += 1
					col = 0

				else :
					for li in lis :
						sheet1.write(row,col,li)
						col += 1
					ncol = col 
					sheet1.write(row,ncol,'计划生产')
					row += 1
					sheet1.write(row,ncol,'计划出货')
					i = 1
					for i in range(31) :
						i += 1
						sheet1.write_formula(row,ncol+i,'IFERROR(VLOOKUP(%s,B:AB,%d+8,0),0)'%(fname,i))
					sheet1.write_formula(row,ncol+32,'SUN(INDIRECT(ADDRESS(%s,%s)):INDIRECT(ADDRESS(%s,%s)))'%(row,ncol,row,ncol+31))
					row += 1
					sheet1.write(row,ncol-1,'库存')
					i = 1
					for i in range(31) :
						i += 1
						sheet1.write_formula(row,ncol+i,'INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s))'%(row+1,ncol+i,row-1,ncol+i+1,row,ncol+i+1))
					sheet1.write_formula(row,ncol+32,'SUM(INDIRECT(ADDRESS(%s,%s)):INDIRECT(ADDRESS(%s,%s)))'%(row,ncol,row,ncol+31))
					row += 1
					col = 0

		f.close()
		print('写入成功，文件保存在：%s'%self.arg)


	"""读取excel文件"""
	def readExcel(self,file):
		book = xlrd.open_workbook(self.arg+file)
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
				j +=1
			if li[1] != '' and li[1] != 'ERP编码' :
				old.append(li)
			
		return data

