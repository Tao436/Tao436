from ntpath import join
import os
import time
import xlrd
import requests
import response

date = time.strftime("%Y%m%d", time.localtime()) #获取当前格式化时间
root = 'E:/用户目录/Desktop/TEST/'#定义工作目录
print (root)
date = root  + date#定义以格式化日期文件夹名称
if not os.path.exists(date):#判断文件夹是否存在
    os.makedirs(date)
xl = root+'图片下载用.xlsx'#工作表路径
xlopen = xlrd.open_workbook(xl)#使用xlrd模块打开工作表
sh_url = xlopen.sheet_by_name('urls')#在工作表内以名称搜索表单
sh_bz = xlopen.sheet_by_name('bz')
urlnrows = sh_url.nrows#获取表单内有效行数
bznrows = sh_bz.nrows
print("图片总数:",urlnrows)
for i in range(bznrows):
    bz = sh_bz.cell_value(i,0)#获取日期文件夹下级文件夹名称-标准号
    folderinbz = date + '/'+bz
    lifolderindl =list(('2D','3D','技术文档'))
    lifolderinyl =list(('1-渲染图','2-工程图','3-线框图','4-3D预览'))
    lipicfolder =list(('1-渲染图','2-工程图','3-线框图'))
    for d in lifolderindl:
        if not os.path.exists(folderinbz+'/'+'下载附件'+'/'+d):
            os.makedirs(folderinbz+'/'+'下载附件'+'/'+d)
    for y in lifolderinyl:
        if not os.path.exists(folderinbz+'/'+'预览附件'+'/'+y):
            os.makedirs(folderinbz+'/'+'预览附件'+'/'+y)
print ('目录创建完成')
folderinyl = folderinbz+'/'+'预览附件'+'/'
yl_1 = str(folderinyl + lipicfolder[0])
yl_2 = str(folderinyl + lipicfolder[1])
yl_3 = str(folderinyl + lipicfolder[2])
liyl = list((yl_1,yl_2,yl_3))
#for p in pics:
    #picstorlocation = (folderinbz+'/'+'预览附件'+'/'+ p +'/')
    #print (picstorlocation)
    #lipiclo = list.insert(picstorlocation)
    #print (lipiclo)
for u in range(urlnrows):
    url = sh_url.cell_value(u,2)#获取图片url    
    picname = sh_url.cell_value(u,0) +'.'+'jpg'#获取图片名称及定义后缀
    #newpic = 'E:\用户目录\Desktop\TEST/'+picname
    newpic = liyl[u%3]+ '/' +picname#定义图片完整路径
    print (newpic)
    req = requests.get(url)#获取url内容
    with open (newpic,'wb')as code:
        code.write(req.content)#保存图片
print('下载完成')   






