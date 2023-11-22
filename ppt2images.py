import win32com
import win32com.client
import sys
import os
import glob
import shutil
from PIL import Image

#获取当前目录
ppt_root = png_root =sys.path[0]+"\\"


def ppt2png(pptFileName):

    powerpoint = win32com.client.Dispatch('PowerPoint.Application')

    #是否后台运行
    powerpoint.Visible = True

    ppt_path = ppt_root + pptFileName

    outputFileName = pptFileName[0:-4] + ".pdf"

    ppt = powerpoint.Presentations.Open(ppt_path)
    #保存为图片
    ppt.SaveAs(png_root + pptFileName.rsplit('.')[0] + '.png', 17) # formatType = 17 ppt转图片
    #保存为pdf
    #ppt.SaveAs(png_root + outputFileName, 32) # formatType = 32 ppt转pdf

    # 关闭打开的ppt文件
    ppt.Close()
    # 关闭powerpoint软件
    powerpoint.Quit()

def pngMontage(dirName):

    #打开目录下所有的png图片
    imageList = []
    file_list = os.listdir(png_root+dirName)

    sorted_file_list = sorted(file_list, key=lambda x: os.path.getctime(os.path.join(png_root+dirName, x)))

    final_file = sorted_file_list
    imageList = final_file
    #获取每张图的宽高
    for file in final_file:
         print(file)
    imageList = [Image.open(png_root+dirName+'\\'+img) for img in final_file  if img.endswith('.JPG')]
    #获取每张图的宽高
    width,height = imageList[0].size
    #新建空白图片并设置图片的宽高,其中高度为所有图片高的总和
    longImage = Image.new(imageList[0].mode,(width*3,int((len(imageList)*height+height*6)/3)))
    begin_x = 0
    begin_y = height*2
    for index,image in enumerate(imageList):
        if (index == 0):
           out = image.resize((width*3,begin_y),Image.ANTIALIAS)
           longImage.paste(out,(begin_x, 0))
           #begin_x += width
        else:
             longImage.paste(image,(begin_x, begin_y))
             begin_x += width
             if begin_x % (width*3) == 0:
                        begin_x = 0
                        begin_y += height
    isExist = os.path.exists(ppt_root+'/预览图')#创建预览图目录
    if not isExist :
        os.mkdir(ppt_root+'/预览图')
    longImage.save(ppt_root+'/预览图'+'/'+dirName+'jpg')#保存生成的图片



#批量打开当前目录下所有的ppt文件
for ppt in (pptFiles for pptFiles in sorted(os.listdir(ppt_root)) if pptFiles.endswith('.pptx') or pptFiles.endswith('.ppt')):
    ppt2png(ppt) #ppt导出图片
    pngMontage(ppt[0:-4]) #所有图片拼接成长图
    shutil.rmtree(png_root+ppt[0:-4])#删除当前生成图片的目录
