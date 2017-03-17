# -*- coding: utf-8 -*-
from math import ceil, floor, fabs
import arcpy.mapping as mapping
from arcpy.sa import *
import arcpy
import os
import time


# 默认分成十类
def dayLegendLabel(start, end, intervals=10):
    L = []

    gap = int((end - start) / intervals)
    if gap < 1:  # 如果首位间隔很小
        gap = 1
        intervals = int(end - start)

    classes = intervals
    not_equal_interval = (end - start) % intervals

    while start <= end:
        L.append(start)
        start += gap
        if not_equal_interval:
            intervals -= 1
            if (end - start) % intervals == 0:
                gap += 1
                break
    while start <= end:
        L.append(start)
        start += gap
    return L, classes


# todo.分类数目需要进一步明确细化，最大不超过11类，最小保存7类
# 默认分类数目自动计算，除非显式给定interval参数
def densityLegendLabel(min, max, intervals=None):
    L = []
    isLast = False  # 是否去除最后一个间隔过小值的标志
    if not intervals:
        if max < 3:  # 值过小的时候默认分成9类
            start = min
            gap = round((max - min) / 9.0, 1)
            if max % gap <= gap / 2 or fabs(max % gap - gap) < 1e-6:
                isLast = True
        elif 3 <= max <= 6.5:
            start = 0
            gap = 0.5
            if max % gap <= 0.25:
                isLast = True
        elif 6.5 < max <= 12:
            start = 0
            gap = 1
            if max % gap <= 0.5:
                isLast = True
        else:
            start = 0
            gap = 2
            if max % gap <= 1:
                isLast = True
    else:
        start = min
        gap = round((max - min) / float(intervals), 1)
        if max % gap <= gap / 2 or fabs(max % gap - gap) < 1e-6:
            isLast = True

    while start < max:
        L.append(start)
        start += gap

    if isLast:
        L.pop()
    L.append(max)
    L = [round(x, 1) for x in L]
    return L, intervals


def mappingProcess(datetime, province, target_area, density_class=None, day_class=10, out_path=None, out_type="TIFF"):
    # todo.控制densityClass和dayClass最多为14
    cwd = os.getcwd()
    infeature = ''.join([cwd, u"/data/LightningBulletin.gdb/", target_area])
    mxd_density_path = ''.join([cwd, u"/data/LightningBulletin.gdb/", target_area, u"闪电密度空间分布模板.mxd"])
    mxd_day_path = ''.join([cwd, u"/data/LightningBulletin.gdb/", target_area, u"地闪雷暴日空间分布模板.mxd"])

    workspace = ''.join([cwd, u"/temp/", province, '/', datetime, '/', target_area, '.gdb'])
    arcpy.env.workspace = workspace
    arcpy.env.overwriteOutput = True

    # *****处理地闪密度地图文档*******
    mxd = mapping.MapDocument(mxd_density_path)
    mask_lyr = mapping.ListLayers(mxd)[2]
    target_lyr = mapping.ListLayers(mxd)[3]
    mask_lyr.replaceDataSource(workspace, "FILEGDB_WORKSPACE", "densityMask")
    target_lyr.replaceDataSource(workspace, "FILEGDB_WORKSPACE", "lightningDensity")
    # 修改图例
    # 计算数据分割值和不显示值
    extract = ExtractByMask("lightningDensity", infeature)
    raster = Raster("lightningDensity")

    class_break_values, density_class = densityLegendLabel(extract.minimum / 100, extract.maximum / 100, density_class)
    excluded_values = ''.join([str(round(raster.minimum - 0.5, 2)), "-", str(round(extract.minimum - 0.5, 2)), ";",
                               str(round(extract.maximum + 0.5, 2)), "-", str(round(raster.maximum + 0.5, 2))])
    sym = target_lyr.symbology
    sym.excludedValues = excluded_values
    sym.classBreakValues = [x * 100 for x in class_break_values]
    # 更改标题和图例
    layout_items = mapping.ListLayoutElements(mxd, "TEXT_ELEMENT")
    layout_items[-1].text = ''.join([datetime, target_area, u"闪电密度空间分布"])  # title标题
    layout_items[-2].text = ''.join([target_area, u"气象局制"])  # 底注

    density_class = class_break_values.__len__() - 1
    class_break_values = [str(x) for x in class_break_values]

    # TODO 将position参数放到设置文件里，或者自动处理
    position = {u"绍兴市": (1.0436, 1.6639, 0.9859272727272728), u"嵊州市": (1.1138, 1.2614, 0.9869090909090698),
                u"上虞区": (1.1372, 1.1617, 0.9862375000000156), u"诸暨市": (1.1138, 1.2614, 0.9869090909090698),
                u"柯桥区": (1.1372, 1.1617, 0.9862375000000156), u"新昌县": (1.1138, 1.2614, 0.9869090909090698),
                u"新乡市": (1.0436, 1.6639, 0.9859272727272728)}

    Y, Xstart, gap = position[target_area]
    for i in xrange(density_class + 1):
        layout_items[i].text = class_break_values[i]
        layout_items[i].elementPositionY = Y
        layout_items[i].elementPositionX = Xstart + gap * i
    for i in xrange(density_class + 1, 15):  # 最多14类,15个标签
        layout_items[i].elementPositionX = -20

    mxd.save()

    if not out_path:
        out_path = workspace
    out_name = os.path.join(out_path, ''.join([datetime, target_area, u"闪电密度空间分布"]))

    out_type = out_type.lower()
    if out_type in ["jpg", 'gpeg']:
        mapping.ExportToJPEG(mxd, out_name, resolution=200)
    elif out_type == "png":
        mapping.ExportToPNG(mxd, out_name, resolution=200)
    elif out_type == "pdf":
        mapping.ExportToPDF(mxd, out_name)
    elif out_type == "bmp":
        mapping.ExportToBMP(mxd, out_name, resolution=200)
    else:
        mapping.ExportToTIFF(mxd, out_name, resolution=200)

    # *****处理雷暴日地图文档*******
    mxd = mapping.MapDocument(mxd_day_path)
    mask_lyr = mapping.ListLayers(mxd)[2]
    target_lyr = mapping.ListLayers(mxd)[3]
    mask_lyr.replaceDataSource(workspace, "FILEGDB_WORKSPACE", "dayMask")
    target_lyr.replaceDataSource(workspace, "FILEGDB_WORKSPACE", "lightningDay")
    # 修改图例
    # 计算数据分割值和不显示值
    extract = ExtractByMask("lightningDay", infeature)
    raster = Raster("lightningDay")
    start = floor(extract.minimum)
    end = ceil(extract.maximum)
    class_break_values, day_class = dayLegendLabel(start, end, day_class)
    excluded_values = ''.join([str(round(raster.minimum - 0.5, 2)), "-", str(round(extract.minimum - 0.5, 2)), ";",
                               str(round(extract.maximum + 0.5, 2)), "-", str(round(raster.maximum + 0.5, 2))])
    sym = target_lyr.symbology
    sym.excludedValues = excluded_values
    sym.classBreakValues = class_break_values
    # 更改标题和图例
    layout_items = mapping.ListLayoutElements(mxd, "TEXT_ELEMENT")
    layout_items[-1].text = ''.join([datetime, target_area, u"地闪雷暴日空间分布"])
    layout_items[-2].text = ''.join([target_area, u"气象局制"])  # 底注

    Y, Xstart, gap = position[target_area]
    for i in xrange(day_class + 1):
        layout_items[i].text = class_break_values[i]
        layout_items[i].elementPositionY = Y
        layout_items[i].elementPositionX = Xstart + gap * i
    for i in xrange(day_class + 1, 15):  # 最多14类,15个标签
        layout_items[i].elementPositionX = -20

    mxd.save()
    out_name = os.path.join(out_path, ''.join([datetime, target_area, u"地闪雷暴日空间分布"]))
    out_type = out_type.lower()
    if out_type in ["jpg", 'gpeg']:
        mapping.ExportToJPEG(mxd, out_name, resolution=200)
    elif out_type == "png":
        mapping.ExportToPNG(mxd, out_name, resolution=200)
    elif out_type == "pdf":
        mapping.ExportToPDF(mxd, out_name)
    elif out_type == "bmp":
        mapping.ExportToBMP(mxd, out_name, resolution=200)
    else:
        mapping.ExportToTIFF(mxd, out_name, resolution=200)


if __name__ == "__main__":
    datetime = u"2015年"
    province = u'浙江'
    target_area = u"绍兴市"

    start = time.clock()
    # ***********************测试程序*********************************"
    mappingProcess(datetime, province, target_area)  # ,density_class, day_class)
    # ***********************测试程序*********************************"
    end = time.clock()
    elapsed = end - start
    print "Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000)
