import os
import re
import requests
import time
import responses
import xlrd

date = time.strftime("%Y%m%d",time.localtime())
root = os.getcwd()
print(root)

xlsx = "图片下载用.xlsx"
xlopen = xlrd.open_workbook(xlsx)
sh = xlopen.sheet_by_name("Sheet1")
bzamount = int(sh.cell_value(rowx=0,colx=0))
col0 = list(sh.col_values(colx=0,start_rowx=1,end_rowx=bzamount))
nrows = int(sh.nrows)

folder_date = root +'/'+ date
print ('图片总数：',nrows)

if not os.path.exists(folder_date):#判断以日期命名的文件夹是否存在
    os.makedirs(folder_date)#不存在即创建
for a in col0 :
    print (a)
    folderindate  = folder_date + '/' + a
    linbz = list(('下载附件','预览附件'))
    lindl = list(('2D','3D','技术文档'))
    linyl = list(('1-渲染图','2-工程图','3-线框图','4-3D预览'))
    linyl_piclocation = list(('1-渲染图','2-工程图','3-线框图'))
    folderdl = folderindate + '/' + linbz[0]
    folderyl = folderindate + '/' + linbz[1]
    for c in lindl:
        folderindl = folderdl +'/' + c
        os.makedirs(folderindl)

    for d in linyl:
        folderinyl = folderyl + '/' +d
        pic_location = folderyl + '/' +d
        yl_1 = folderyl + '/'+ linyl[0]
        yl_2 = folderyl + '/'+ linyl[1]
        yl_3 = folderyl + '/'+ linyl[2]
        yl_4 = folderyl + '/'+ linyl[3]
        os.makedirs(folderinyl)
for i in range(0,nrows):
    pic_url = sh.cell_value(rowx=i,colx=4)
    pic_name = sh.cell_value(rowx=i,colx=2)+'.'+'jpg'
    req = requests.get(pic_url)
    pic_location = folderinyl = folderyl + '/' + d + '/'
    new_pic = pic_location + pic_name
    with open (new_pic,'wb') as code:
        code.write(req.content)
print('sucess')

    


'''print("图片总数",nrows)
for i in range(nrows):
    print("下载第",i+1,"个图片")
    name=sheet.cell_value(i,0)#读取图片名称
    url = sheet.cell_value(i,2)  # 依次读取每行第3列的数据，也就是 URL
    f=requests.get(url)#下载图片
    pic_location = "E:/用户目录/Desktop/图片下载程序/"+ time_now +"/"
    pic_name = name + "." + "jpg"  # 构造完整文件路径+名称
    newpic =  pic_location + pic_name
    with open(newpic,'wb') as code:
         code.write(f.content)  # 保存文件
print("下载完成")'''
            
    
    



"""li1 = ('预览附件','下载附件')
li2_1 = ('2D','3D','技术文档')
li2_2 = ('1-渲染图','2-工程图','3-线框图','4-3D预览')

path1 = root +'\\' + date+'\\'+ bz
path2_1 = path1 +'\\' + li1[0]
path2_2 = path1 +'\\' + li1[1]
path3_1 = path1 + ''

if not os.path.exists(path1):
    os.makedirs(path1)
    os.makedirs(path2_1)
    os.makedirs(path2_2)


print (path2_2)"""

