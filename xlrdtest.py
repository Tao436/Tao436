import xlrd
import os

from xlrd import sheet 
root = os.getcwd()
xl = '图片下载用.xlsx'
xlopen = xlrd.open_workbook(xl)
sh = xlopen.sheet_by_name('Sheet1')
bzamount = sh.cell_value(rowx=0,colx=0)
bzv = int(bzamount)
nrows = sh.nrows
col4 = sh.col_values(colx=3,start_rowx=0,end_rowx=bzv)

for i in range( 0,nrows ):
        while i<= nrows:
            pic_name = sh.cell_value(1,i)
            
            print (pic_name)

#print(type (nrows))

