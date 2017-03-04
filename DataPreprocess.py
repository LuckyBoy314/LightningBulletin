# -*- coding: utf-8 -*-
#import tkinter.filedialog as tkFileDialog
import tkFileDialog
import time
#import tkinter as Tkinter
import Tkinter
import arcpy
import os
from arcpy.da import SearchCursor
import cPickle as pickle

"""
主要功能：
    1.writeOrignData
    将原始电闪数据（txt格式，每天一个txt）生成单一的txt文档，并符合ArcGIS表格数据格式要求
    2.generateDatabase
    将第一步生成的txt表格数据与省级分县地图叠加，为各个闪电定位点生成省-市-县三个字段信息
    生成的数据为点要素类，存储在文件数据库中，这是后续地理处理的基础数据；
    同时在主调函数preProcess中将点要素类转换为个人数据库格式，这是后续查询的基础

    注意：无论是文件数据库还是个人数据库，其中的要素类其实是一样的，都是以省为单位的一年闪电定位数据，
    只是格式不一样，文件数据库速度更快，个人数据库速度慢一点，而且有2G容量限制，但可以作为SQL查询。

主要输入：
    1.原始文件路径
    2.时间（年份）
    3.省份
主要输出：
    1. 矢量数据database，记录各个电闪定位数据的物理参数和地理位置信息，是后续处理的数据基础
    2. SQL文件，通过将矢量数据database转换为个人数据，生产ACCESS文件，供后续查询
中间文件：
    1.合成txt表格数据
    2.origindata_withOID.shp临时文件
"""

def writeOrignData(infiles, text_table):
    if os.path.isfile(text_table):
        return
    with open(text_table, 'w') as out_f:
        out_f.write('Date,Time,Lat,Lon,Intensity,Slope,Error,LocateStyle\n')
        for infile in infiles:
            with open(infile, 'r') as in_f:
                for line in in_f:
                    line = line.decode('gb2312').split()
                    line[3:8] = [x[3:] for x in line[3:8]]
                    line[8] = line[8][5:]
                    s = u','.join(line[1:9])
                    s = u''.join([s,'\n'])
                    out_f.write(s.encode('utf-8'))


def dialogOpenFile():
    master = Tkinter.Tk()
    master.withdraw()  # 不显示tk主窗口

    files = tkFileDialog.askopenfilenames(  # 调用系统窗口选择文件
            initialdir=u'Z:/ZHUF/数据/闪电数据',
            title=u'请选择原始的电闪数据',
            filetypes=[("text file", "*.txt"), ("all", "*.*")]

    )
    return files


def FTPOpenFile():
    pass


def generateDatabase(text_table, province):
    # 输出文件名根据输入名确定
    database = text_table.replace(".txt", "")
    if arcpy.Exists(database):
        return database

    # Make the XY event
    originData = "originData"
    arcpy.MakeXYEventLayer_management(text_table, "Lon", "Lat", originData,
                                      r"Coordinate Systems\Geographic Coordinate Systems\World\WGS 1984")

    # XY event 没有OID，有些操作需要OID，如intersect，有的不需要，如clip
    originData_withOID = os.path.join(os.path.dirname(text_table), "origindata_withOID")
    #if arcpy.Exists(originData_withOID):
    #    arcpy.Delete_management(originData_withOID)
    arcpy.CopyFeatures_management(originData, originData_withOID)

    cwd = os.getcwd()
    province_feature= ''.join([cwd,u'/data/LightningBulletin.gdb/', province,u'_分县'])
    arcpy.Intersect_analysis([originData_withOID, province_feature], database)

    arcpy.DeleteField_management(database, ["FID_origindata_withOID", u"FID_"+ province, 'Code', 'Shape_Leng', 'Shape_Area'])  # 清理不必要的字段
    arcpy.Delete_management(originData_withOID)#删除中间生成的辅助临时文件
    return database

# 同一年的数据 预处理只需要做一次就行了,这个处理应该放在模块函数中解决，而不是放在主模块中
def preProcess(datetime, province,infiles):
    # 根据datetime创立工作目录
    cwd = os.getcwd()
    workpath = ''.join([cwd,u"/temp/",province,'/', datetime])
    if not os.path.exists(workpath):
        os.makedirs(workpath)
    workspace = ''.join([workpath,'/GDB.gdb'])
    if not arcpy.Exists(workspace):
        arcpy.CreateFileGDB_management(workpath,'GDB.gdb')
    arcpy.env.overwriteOutput = True

    text_table = ''.join([workspace , "/" , u"data" , datetime , u".txt"])
    writeOrignData(infiles, text_table)
    database_path = generateDatabase(text_table, province)

    #建立数据库,以供SQL查询
    if not arcpy.Exists(''.join([workpath,'/SQL.mdb'])):
        arcpy.CreatePersonalGDB_management(workpath, "SQL.mdb")#在指定位置建立个人数据库
        arcpy.FeatureClassToGeodatabase_conversion(database_path,''.join([workpath,'/SQL.mdb']))#将文件数据库中的要素类导入到个人数据库

    #计算省下属各地级市面积
    province_area = {}
    province_feature = ''.join([cwd,u'/data/LightningBulletin.gdb/', province,u'_分区'])
    with SearchCursor(province_feature, ["Region","area"]) as cursor:
        for row in cursor:
            province_area[row[0]] = row[1]

    f = file(os.path.join(workspace,'province_area.pkl'), 'wb')
    pickle.dump(province_area, f, pickle.HIGHEST_PROTOCOL)
    f.close()

if __name__ == "__main__":
    datetime = u"2016年"
    province = u'河南'
    infiles = dialogOpenFile()

    start = time.clock()
    # ***********************测试程序*********************************"
    print(preProcess(datetime,province,infiles))
    # ***********************测试程序*********************************"
    end = time.clock()
    elapsed = end - start
    print("Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000))
