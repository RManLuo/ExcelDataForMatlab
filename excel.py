# -*- coding: utf-8 -*-
import xlrd
import xlwt
import xlutils
import sys
import xlsxwriter

def main():
	workbook=xlrd.open_workbook(sys.argv[1])
	workbook_out=xlsxwriter.Workbook(sys.argv[1]+'_out.xlsx')
	sheet_out=workbook_out.add_worksheet('sheet1')
	sheet=workbook.sheets()[0]
	nrows=sheet.nrows
	nclos=sheet.ncols
	for i in range(0,nclos):
		d=0
		for j in sheet.col_values(i):
			d=d+1
			
			now=sheet.col(i)[0].value
			
			
			if  isinstance(j,float):
				sheet_out.write(d,i,j)
				continue
			if ((not isinstance(j, float)) and j == now):
				sheet_out.write(d,i,j)
				continue
			if 'u' in j:
				
				x=float(j.split('u')[0])/1000000
				sheet_out.write(d,i,x)
			if 'm' in j:
				
				x=float(j.split('m')[0])/1000
				sheet_out.write(d,i,x)
	workbook_out.close()
if __name__=='__main__':
	print '''
	Automaticly change 'm' and 'u' into float
	Useage: in cmd execute execl.exe filename.xls 
	OUTPUT: filename.xls_out.xlsx
	'''
	if len(sys.argv) <= 1:
		print "Please input filename"
		os._exit(0)
	main()
		
	