import requests
from requests.models import codes
import xlrd

xl = '图片下载用.xlsx'
xo = xlrd.open_workbook(xl)
sh = xo.sheet_by_name('urls')
nrows =sh.nrows
#listl = sheet.col_values(colx=0,start_rowx=0,end_rowx=None)
#print (listl)
url=sh.cell_value(0,2)
print(url)
with open ('E:\用户目录\Desktop\TEST/20210809\ISO14581\预览附件/1-渲染图/'+'a'+'.'+'jpg','wb') as code:   
    code.write(requests.get(url).content)
print ('sucess')



