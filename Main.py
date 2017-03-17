# -*- coding: utf-8 -*-
from DataPreprocess import preProcess
from GeoProcess import geoProcess
from Mapping import mappingProcess
import tkFileDialog
import Tkinter
import subprocess
import os

"""
主要输入：
    origin_data_path:原始数据路径：
    target_area：目标区域
    datetime:时间，年度公报 月报 日报 任意时间间隔
    density_cell, density_class:电闪密度网格大小，图例分类个数
    day_cell, day_class：地闪雷暴日网格大小，图例分类个数
    out_type:输出格式
    out_path:输出路径
主要输出：
    图片(相应格式）
    统计信息
"""

def dialogOpenFile():
    master = Tkinter.Tk()
    master.withdraw()  # 不显示tk主窗口

    files = tkFileDialog.askopenfilenames(  # 调用系统窗口选择文件
            initialdir=u'Z:/ZHUF/数据/闪电数据',
            title=u'请选择原始的电闪数据',
            filetypes=[("text file", "*.txt"), ("all", "*.*")]

    )
    return files

def mainProcess(datetime, province,target_area, origin_data_path,
                density_cell = '10', day_cell = '15',density_class = None, day_class = 10,
                out_type = 'TIFF', out_path=None):

    cwd = os.getcwd() #注意cwd在程序执行过程中会变化

    #数据预处理
    preProcess(datetime, province, origin_data_path)

    # 地理处理，生成作图文件
    geoProcess(datetime, province, target_area, density_cell, day_cell)

    # 地图处理
    """
    暂时采用默认参数
    density_class =
    day_class =
    out_type =
    out_path =
    """
    mappingProcess(datetime, province, target_area)

    #启动外部程序，调用python3处理，注意命令行编码是gbk
    subprocess.call([u'C:/Program Files/Python35/python.exe'.encode('gbk'),
                     (cwd + u'/StatsProcess.py').encode('gbk'),
                     datetime.encode('gbk'),province.encode('gbk'),target_area.encode('gbk'),cwd.encode('gbk')])

if __name__ == "__main__":
    datetime = u"2015年"
    province = u'浙江'
    target_area = u"绍兴市"
    infiles = dialogOpenFile()

    start = time.clock()
    # ***********************测试程序*********************************"
    mainProcess(datetime, province, target_area, infiles)
    # ***********************测试程序*********************************"
    end = time.clock()
    elapsed = end - start
    print("Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000))
