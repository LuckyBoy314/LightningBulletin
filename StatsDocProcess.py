# -*- coding: utf-8-*-

from win32com.client import DispatchEx, constants
from win32com.client.gencache import EnsureDispatch
import pyodbc
import os
import time
import pickle
import sys
import shutil


def dict2list(dic: dict):
    ''' 将字典转化为列表 '''
    keys = dic.keys()
    vals = dic.values()
    lst = [(key, val) for key, val in zip(keys, vals)]
    return lst


def sqlQuery(workbook, year, province, target_area, cwd):
    local_path = ''.join([cwd, u"/temp/", province, '/', year, '/', target_area, '.gdb'])
    query_results = {}

    # 链接数据库
    database = ''.join([cwd, u"/temp/", province, '/', year, '/SQL.mdb;'])

    db = pyodbc.connect(''.join(['DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};', 'DBQ=', database]))
    cursor = db.cursor()
    data_table = ''.join(['data', year])  # 查询表

    # 读取省下属各地级市面积
    workspath = ''.join([cwd, u"/temp/", province, '/', year, '/', 'GDB.gdb'])
    f = open(os.path.join(workspath, 'province_area.pkl'), 'rb')
    province_area = pickle.load(f)  # 诸如{'新乡市': XXX, '安阳市': XXX}
    f.close()

    workspath = ''.join([cwd, u"/temp/", province, '/', year, '/', target_area, '.gdb'])
    # 读取地级市下属各县市面积
    f = open(os.path.join(workspath, 'region_area.pkl'), 'rb')
    region_area = pickle.load(f)  # 诸如{'新乡市'：XXX, '延津县':XXX}
    f.close()

    # 读取目标地级区域的雷暴日和闪电密度的统计信息
    f = open(os.path.join(workspath, 'stats.pkl'), 'rb')
    stat_density = pickle.load(f)  # 诸如{'新乡':(最大值，最小值)}
    stat_day = pickle.load(f)  # 诸如{'新乡':(平均值，最大值，最小值)}
    f.close()

    try:
        # ************分地区统计**********
        sql = """
        SELECT count(*) AS num, Region
        FROM %s
        WHERE Province= \'%s\'
        GROUP BY Region
        ORDER BY count(*) DESC
        """ % (data_table, province + '省')

        # 处理SQL查询结果，顺便记录本地区地闪次数在全省的排名
        sum_province_dict = {}  # 全省各地区地闪总数
        rank = 0
        for row in cursor.execute(sql):
            sum_province_dict[row[1]] = row[0]  # 以 Region：num建立字典，方便下面赋值
            rank += 1
            # print type(row[1])
            if row[1] == target_area:
                Sum_rank_in_province = rank  # 本地区地闪次数在全省的排名

        Sum_target_area = sum_province_dict[target_area]  # 本地区地闪总数
        Density_target_area = round(Sum_target_area / province_area[target_area], 2)  # 本地区地闪密度

        # 计算全省各地区地闪密度和全省平均密度
        density_province_dict = {}  # 全省各地区密度
        Density_province = 0
        for key in sum_province_dict:
            density_province_dict[key] = round(sum_province_dict[key] / province_area[key], 2)
            Density_province += density_province_dict[key]
        Density_province = round(Density_province / len(province_area), 2)  # 全省平均地闪密度
        # 密度从大到小进行排序
        density_province_sorted = sorted(dict2list(density_province_dict), key=lambda d: d[1], reverse=True)

        # 计算本地区地闪密度排名
        rank = 0
        for item in density_province_sorted:
            rank += 1
            if item[0] == target_area:
                Density_rank_in_province = rank  # 本地区地闪密度在全省的排名
                break

        # 省分地区的电闪次数和密度
        Stats_of_Province = {}
        sum_of_province = 0
        for key in sum_province_dict:
            sum_of_province += sum_province_dict[key]
            Stats_of_Province[key[:-1]] = (sum_province_dict[key], density_province_dict[key])
        Stats_of_Province[u'总计'] = (sum_of_province, Density_province)

        query_results['Sum_target_area'] = Sum_target_area  # 本地区地闪总数
        query_results['Sum_rank_in_province'] = Sum_rank_in_province  # 本地区地闪总数在全省排名
        query_results['Density_province'] = Density_province  # 全省密度
        query_results['Density_target_area'] = Density_target_area  # 本地区密度
        query_results['Density_rank_in_province'] = Density_rank_in_province  # 本地区密度在全省排名
        query_results['Stats_of_Province'] = Stats_of_Province  # 省分地区统计

        # ********* 分县统计***********
        sql = """
        SELECT count(*) AS num, County
        FROM %s
        WHERE Region=\'%s\'
        GROUP BY County
        ORDER BY count(*) DESC
        """ % (data_table, target_area)

        # 处理SQL查询结果，顺便记录地闪次数的最大和最小值
        sum_region_dict = {}  # 本地区下属各县地闪总数
        rank = 0
        num_region = len(region_area)
        for row in cursor.execute(sql):
            sum_region_dict[row[1]] = row[0]  # 以 Region：num建立字典，方便下面赋值
            rank += 1
            if rank == 1:
                Sum_max_county_name = row[1]  # 闪电次数最多的县名
                Sum_max_in_region = row[0]  # 以及次数
            elif rank == num_region:
                Sum_min_county_name = row[1]  # 闪电次数最少的县名
                Sum_min_in_region = row[0]  # 以及次数

        # 计算地闪总数最大、最小的县的占目标区域总数的比例
        Max_region_percent = round(Sum_max_in_region / float(Sum_target_area) * 100, 2)
        Min_region_percent = round(Sum_min_in_region / float(Sum_target_area) * 100, 2)

        # 计算本地区各县市地闪密度
        density_region_dict = {}

        for key in sum_region_dict:
            density_region_dict[key] = round(sum_region_dict[key] / region_area[key], 2)
        # 密度从大到小进行排序
        density_region_sorted = sorted(dict2list(density_region_dict), key=lambda d: d[1], reverse=True)
        # 最大、最小密度
        Density_max_county_name = density_region_sorted[0][0]  # 闪电密度最高的县名
        Density_max_in_region = density_region_sorted[0][1]  # 闪电密度最高的县密度
        Density_min_county_name = density_region_sorted[num_region - 1][0]  # 闪电密度最低的县名
        Density_min_in_region = density_region_sorted[num_region - 1][1]  # 闪电密度最低的县密度

        Stats_of_Region = {}  # 地闪次数、平均地闪密度、最大地闪密度、最小地闪密度、平均雷暴日、最大雷暴日、最小雷暴日
        for key in sum_region_dict:
            Stats_of_Region[key] = (sum_region_dict[key], density_region_dict[key],
                                    stat_density[key[:-1]][0], stat_density[key[:-1]][1],
                                    stat_day[key[:-1]][0], stat_day[key[:-1]][1], stat_day[key[:-1]][2])
        max_density_region = max([stat_density[each][0] for each in stat_density])
        min_density_region = min([stat_density[each][1] for each in stat_density])
        Day_target_area = int(sum([stat_day[each][0] for each in stat_day]) / num_region)
        max_day_region = max([stat_day[each][1] for each in stat_day])
        min_day_region = min([stat_day[each][2] for each in stat_day])

        Stats_of_Region[u'总计'] = (Sum_target_area, Density_target_area, max_density_region, min_density_region,
                                  Day_target_area, max_day_region, min_day_region)

        query_results['Sum_max_county_name'] = Sum_max_county_name  # 闪电次数最多的县名
        query_results['Sum_max_in_region'] = Sum_max_in_region  # 闪电次数最多的县次数
        query_results['Sum_min_county_name'] = Sum_min_county_name  # 闪电次数最少的县名
        query_results['Sum_min_in_region'] = Sum_min_in_region  # 闪电次数最少的县次数
        query_results['Density_max_county_name'] = Density_max_county_name  # 闪电密度最多的县名
        query_results['Density_max_in_region'] = Density_max_in_region  # 闪电密度最多的县密度
        query_results['Density_min_county_name'] = Density_min_county_name  # 闪电密度最少的县名
        query_results['Density_min_in_region'] = Density_min_in_region  # 闪电密度最少的县密度
        query_results['Max_region_percent'] = Max_region_percent  # 地闪次数最大县所占比例
        query_results['Min_region_percent'] = Min_region_percent  # 地闪次数最小县所占比例
        query_results['Stats_of_Region'] = Stats_of_Region  # 分县市统计

        # 分区统计结果写入excel
        sheet = workbook.Worksheets(u'分区统计')
        if province == u'浙江':
            regions = [u'杭州', u'宁波', u'湖州', u'嘉兴', u'绍兴', u'金华', u'台州', u'温州', u'衢州', u'丽水', u'舟山', u'总计']
            counties = [u'越城区', u'柯桥区', u'上虞区', u'诸暨市', u'嵊州市', u'新昌县', u'总计']
        elif province == u'河南':
            regions = [u'郑州', u'开封', u'洛阳', u'平顶山', u'安阳', u'鹤壁', u'新乡', u'焦作', u'濮阳',
                       u'许昌', u'漯河', u'三门峡', u'商丘', u'周口', u'驻马店', u'南阳', u'信阳', u'济源', u'总计']
            counties = [u'市区', u'新乡县', u'辉县市', u'卫辉市', u'获嘉县', u'原阳县', u'延津县', u'封丘县', u'长垣县', u'总计']

        n_regions = len(regions)
        for i in range(n_regions):
            sheet.Cells(1, i + 2).Value = regions[i]
            sheet.Cells(2, i + 2).Value = Stats_of_Province[regions[i]][0]
            sheet.Cells(3, i + 2).Value = Stats_of_Province[regions[i]][1]

        n_counties = len(counties)
        for i in range(n_counties):  # i表示列，一个区域的序号
            sheet.Cells(5, i + 2).Value = counties[i]
            for j in range(6, 13):  # j表示行
                sheet.Cells(j, i + 2).Value = Stats_of_Region[counties[i]][j - 6]

        # todo SQL查询有待优化
        # ************分月统计 月地闪次数和月平均强度(负闪)**************
        sql = """
        SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,1 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,2 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,3 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,4 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,5 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,6 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,7 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,8 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,9 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,10 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,11 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
        UNION SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,12 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
        ORDER BY 月份
        """.replace('TARGET_AREA', target_area).replace('YEAR', year[0:4]).replace('QUERY_TABLE', data_table)

        sheet = workbook.Worksheets(u'分月统计')
        Stats_of_Month = {}
        i = 1  # 行号
        month = 0
        sum_month_dict = {}  # 记录分月地闪总数
        negative_intensity_dict = {}  # 记录分月负闪地闪次数
        for row in cursor.execute(sql):
            i += 1
            month += 1
            sheet.Cells(i, 2).Value = row[0]  # 负闪次数
            sheet.Cells(i, 5).Value = negative_intensity_dict[month] = row[1] if row[1] is not None else 0  # 负闪强度
            sum_month_dict[month] = row[0]
            Stats_of_Month[month] = [row[0], negative_intensity_dict[month]]

        # 负闪强度峰值所在月份
        negative_intensity_sorted = sorted(dict2list(negative_intensity_dict), key=lambda d: d[1], reverse=True)
        Peak_month_negative_intensity = negative_intensity_sorted[0][0]

        # ************分月统计 月地闪次数和月平均强度(正闪)**************
        sql = """
        SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,1 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,2 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,3 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,4 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,5 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,6 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,7 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,8 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,9 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,10 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,11 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,12 AS 月份
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
        ORDER BY 月份
        """.replace('TARGET_AREA', target_area).replace('YEAR', year[0:4]).replace('QUERY_TABLE', data_table)

        i = 1  # 行号
        month = 0
        positive_intensity_dict = {}  # 记录正闪强度
        for row in cursor.execute(sql):
            i += 1
            month += 1
            sheet.Cells(i, 3).Value = row[0]  # 正闪次数
            sheet.Cells(i, 6).Value = positive_intensity_dict[month] = row[1] if row[1] is not None else 0  # 正闪强度
            sum_month_dict[month] += row[0]
            Stats_of_Month[month].append(row[0])
            Stats_of_Month[month].append(positive_intensity_dict[month])
            Stats_of_Month[month].append(sum_month_dict[month])

        # 修改标题
        title = sheet.ChartObjects(1).Chart.ChartTitle.Text
        sheet.ChartObjects(1).Chart.ChartTitle.Text = title.replace(u'2015年', year).replace(u'绍兴市', target_area)
        # 导出分月统计图
        sheet.ChartObjects(1).Chart.Export(''.join([local_path, '/', 'month_stats_pic.png']))

        # 正闪峰值月份
        positive_intensity_sorted = sorted(dict2list(positive_intensity_dict), key=lambda d: d[1], reverse=True)
        Peak_month_positive_intensity = positive_intensity_sorted[0][0]

        sum_month_sorted = sorted(dict2list(sum_month_dict), key=lambda d: d[1], reverse=True)
        # 地闪次数最多的月份
        Max_month_region = sum_month_sorted[0][0]
        # 地闪次数最多的三个月
        Max_three_months = [sum_month_sorted[0][0], sum_month_sorted[1][0], sum_month_sorted[2][0]]
        Max_three_months.sort()
        # 地闪次数最多三个月所占比例
        Max_months_percent = round(
            100 * (sum_month_sorted[0][1] + sum_month_sorted[1][1] + sum_month_sorted[2][1]) / float(
                Sum_target_area), 2)
        # 没有检测到地闪的月份
        Months_zero = [i[0] for i in sum_month_sorted if i[1] == 0]
        Months_zero.sort()

        query_results['Peak_month_negative_intensity'] = Peak_month_negative_intensity  # 正闪峰值月份
        query_results['Peak_month_positive_intensity'] = Peak_month_positive_intensity  # 负闪峰值月份
        query_results['Max_month_region'] = Max_month_region  # 地闪次数最多的月份
        query_results['Max_three_months'] = Max_three_months  # 地闪次数最多的三个月
        query_results['Max_months_percent'] = Max_months_percent  # 地闪次数最多三个月所占比例
        query_results['Months_zero'] = Months_zero  # 没有检测到地闪的月份
        query_results['Stats_of_Month'] = Stats_of_Month  # 分月统计，负闪次数、负闪强度、正闪次数、正闪强度、总次数

        # ****查询雷暴初日********
        sql = """SELECT TOP 1 Date_
        From %s
        Where Region = \'%s\'
        Order By Date_, OBJECTID
        """ % (data_table, target_area)

        for row in cursor.execute(sql):
            First_date = row[0].strftime('%mM%dD').replace('M', '月').replace('D', '日')

        query_results['First_date'] = First_date  # 雷暴初日

        # ************分时段统计 时段地闪次数和时段平均强度(负闪)**************
        sql = """
        SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,0 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=0
        UNION SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,1 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=1
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,2 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=2
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,3 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=3
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,4 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=4
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,5 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=5
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,6 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=6
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,7 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=7
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,8 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=8
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,9 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=9
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,10 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=10
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,11 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=11
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,12 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=12
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,13 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=13
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,14 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=14
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,15 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=15
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,16 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=16
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,17 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=17
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,18 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=18
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,19 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=19
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,20 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=20
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,21 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=21
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,22 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=22
        union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,23 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=23
        ORDER BY 时段
        """.replace('TARGET_AREA', target_area).replace('QUERY_TABLE', data_table)

        sheet = workbook.Worksheets(u'分时段统计')
        sum_hour_dict = {}  # 记录分时段地闪总数
        Stats_of_Hour = {}
        i = 1  # 行号
        for row in cursor.execute(sql):
            i += 1
            hour = i - 2
            sheet.Cells(i, 2).Value = row[0]  # 负闪次数
            sheet.Cells(i, 5).Value = negative_intensity_hours = row[1] if row[1] is not None else 0  # 负闪强度
            sum_hour_dict[hour] = row[0]
            Stats_of_Hour[hour] = [row[0], negative_intensity_hours]

        # ************分时段统计 时段地闪次数和时段平均强度(正闪)**************
        sql = """
        SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,0 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=0
        UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,1 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=1
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,2 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=2
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,3 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=3
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,4 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=4
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,5 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=5
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,6 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=6
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,7 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=7
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,8 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=8
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,9 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=9
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,10 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=10
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,11 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=11
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,12 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=12
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,13 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=13
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,14 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=14
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,15 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=15
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,16 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=16
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,17 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=17
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,18 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=18
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,19 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=19
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,20 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=20
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,21 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=21
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,22 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=22
        union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,23 AS 时段
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=23
        ORDER BY 时段
        """.replace('TARGET_AREA', target_area).replace('QUERY_TABLE', data_table)

        i = 1  # 行号
        for row in cursor.execute(sql):
            i += 1
            hour = i - 2
            sheet.Cells(i, 3).Value = row[0]  # 正闪次数
            sheet.Cells(i, 6).Value = positive_intensity_hours = row[1] if row[1] is not None else 0  # 正闪强度
            sum_hour_dict[hour] += row[0]
            Stats_of_Hour[hour].append(row[0])
            Stats_of_Hour[hour].append(positive_intensity_hours)
            Stats_of_Hour[hour].append(sum_hour_dict[hour])

        query_results['Stats_of_Hour'] = Stats_of_Hour

        # 修改标题
        title = sheet.ChartObjects(1).Chart.ChartTitle.Text
        sheet.ChartObjects(1).Chart.ChartTitle.Text = title.replace(u'2015年', year).replace(u'绍兴市', target_area)
        # 导出分时段统计图
        sheet.ChartObjects(1).Chart.Export(''.join([local_path, '/', 'hour_stats_pic.png']))

        # **********负闪强度分布**************
        sql = """
        SELECT count(*) AS 负闪次数,0 AS 左边界,5 AS 右边界
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=0 AND Abs(Intensity)<5
        union SELECT count(*) AS 负闪次数,5,10
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=5 AND Abs(Intensity)<10
        union SELECT count(*) AS 负闪次数,10,15
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=10 AND Abs(Intensity)<15
        union SELECT count(*) AS 负闪次数,15,20
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=15 AND Abs(Intensity)<20
        union SELECT count(*) AS 负闪次数,20,25
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=20 AND Abs(Intensity)<25
        union SELECT count(*) AS 负闪次数,25,30
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=25 AND Abs(Intensity)<30
        union SELECT count(*) AS 负闪次数,30,35
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=30 AND Abs(Intensity)<35
        union SELECT count(*) AS 负闪次数,35,40
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=35 AND Abs(Intensity)<40
        union SELECT count(*) AS 负闪次数,40,45
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=40 AND Abs(Intensity)<45
        union SELECT count(*) AS 负闪次数,45,50
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=45 AND Abs(Intensity)<50
        union SELECT count(*) AS 负闪次数,50,55
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=50 AND Abs(Intensity)<55
        union SELECT count(*) AS 负闪次数,55,60
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=55 AND Abs(Intensity)<60
        union SELECT count(*) AS 负闪次数,60,65
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=60 AND Abs(Intensity)<65
        union SELECT count(*) AS 负闪次数,65,70
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=65 AND Abs(Intensity)<70
        union SELECT count(*) AS 负闪次数,70,75
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=70 AND Abs(Intensity)<75
        union SELECT count(*) AS 负闪次数,75,80
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=75 AND Abs(Intensity)<80
        union SELECT count(*) AS 负闪次数,80,85
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=80 AND Abs(Intensity)<85
        union SELECT count(*) AS 负闪次数,85,90
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=85 AND Abs(Intensity)<90
        union SELECT count(*) AS 负闪次数,90,95
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=90 AND Abs(Intensity)<95
        union SELECT count(*) AS 负闪次数,95,100
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=95 AND Abs(Intensity)<100
        union SELECT count(*) AS 负闪次数,100,150
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=100 AND Abs(Intensity)<150
        union SELECT count(*) AS 负闪次数,150,200
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=150 AND Abs(Intensity)<200
        union SELECT count(*) AS 负闪次数,200,250
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=200 AND Abs(Intensity)<250
        union SELECT count(*) AS 负闪次数,250,300
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=250 AND Abs(Intensity)<300
        UNION SELECT count(*) AS 负闪次数,300,1000
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=300
        ORDER BY 左边界
        """.replace('TARGET_AREA', target_area).replace("QUERY_TABLE", data_table)

        Stats_of_Intensity = {}
        sheet = workbook.Worksheets(u'强度分布统计')
        i = 1  # 行号
        for row in cursor.execute(sql):
            i += 1
            sheet.Cells(i, 3).Value = row[0]  # 负闪次数
            Stats_of_Intensity[i - 2] = [row[0]]

        # ***********正闪强度分布************
        sql = """
        SELECT count(*) AS 正闪次数,0 AS 左边界,5 AS 右边界
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Intensity<5
        union SELECT count(*) AS 正闪次数,5,10
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=5 AND Intensity<10
        union SELECT count(*) AS 正闪次数,10,15
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=10 AND Intensity<15
        union SELECT count(*) AS 正闪次数,15,20
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=15 AND Intensity<20
        union SELECT count(*) AS 正闪次数,20,25
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=20 AND Intensity<25
        union SELECT count(*) AS 正闪次数,25,30
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=25 AND Intensity<30
        union SELECT count(*) AS 正闪次数,30,35
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=30 AND Intensity<35
        union SELECT count(*) AS 正闪次数,35,40
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=35 AND Intensity<40
        union SELECT count(*) AS 正闪次数,40,45
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=40 AND Intensity<45
        union SELECT count(*) AS 正闪次数,45,50
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=45 AND Intensity<50
        union SELECT count(*) AS 正闪次数,50,55
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=50 AND Intensity<55
        union SELECT count(*) AS 正闪次数,55,60
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=55 AND Intensity<60
        union SELECT count(*) AS 正闪次数,60,65
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=60 AND Intensity<65
        union SELECT count(*) AS 正闪次数,65,70
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=65 AND Intensity<70
        union SELECT count(*) AS 正闪次数,70,75
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=70 AND Intensity<75
        union SELECT count(*) AS 正闪次数,75,80
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=75 AND Intensity<80
        union SELECT count(*) AS 正闪次数,80,85
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=80 AND Intensity<85
        union SELECT count(*) AS 正闪次数,85,90
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=85 AND Intensity<90
        union SELECT count(*) AS 正闪次数,90,95
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=90 AND Intensity<95
        union SELECT count(*) AS 正闪次数,95,100
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=95 AND Intensity<100
        union SELECT count(*) AS 正闪次数,100,150
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=100 AND Intensity<150
        union SELECT count(*) AS 正闪次数,150,200
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=150 AND Intensity<200
        union SELECT count(*) AS 正闪次数,200,250
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=200 AND Intensity<250
        union SELECT count(*) AS 正闪次数,250,300
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=250 AND Intensity<300
        UNION SELECT count(*) AS 正闪次数,300,1000
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=300
        ORDER BY 左边界
        """.replace('TARGET_AREA', target_area).replace("QUERY_TABLE", data_table)

        i = 1  # 行号
        for row in cursor.execute(sql):
            i += 1
            sheet.Cells(i, 4).Value = row[0]  # 正闪次数
            Stats_of_Intensity[i - 2].append(row[0])

        for i in range(25):
            Stats_of_Intensity[i].append(sheet.Cells(i + 2, 6).Value)
            Stats_of_Intensity[i].append(sheet.Cells(i + 2, 7).Value)

        query_results['Stats_of_Intensity'] = Stats_of_Intensity
        # 修改标题
        title = sheet.ChartObjects(1).Chart.ChartTitle.Text
        sheet.ChartObjects(1).Chart.ChartTitle.Text = title.replace(u'2015年', year).replace(u'绍兴市', target_area)
        title = sheet.ChartObjects(2).Chart.ChartTitle.Text
        sheet.ChartObjects(2).Chart.ChartTitle.Text = title.replace(u'2015年', year).replace(u'绍兴市', target_area)
        # 导出分强度统计图
        sheet.ChartObjects(1).Chart.Export(''.join([local_path, '/', 'negative_stats_pic.png']))
        sheet.ChartObjects(2).Chart.Export(''.join([local_path, '/', 'positive_stats_pic.png']))

        f = open(os.path.join(workspath, 'query_results.pkl'), 'wb')
        pickle.dump(query_results, f, 2)  # 不能采用最高的协议，否则在python2.7.8中无法加载进来
        f.close()

    finally:
        db.close()  # 关闭数据连接

    return query_results


def docProcess(word, doc, workbook, query_results, year, province, target_area, cwd):
    sum_target_area = query_results['Sum_target_area']  # 本地区地闪总数
    sum_rank_in_province = query_results['Sum_rank_in_province']  # 本地区地闪总数在全省排名
    density_province = query_results['Density_province']  # 全省密度
    density_target_area = query_results['Density_target_area']  # 本地区密度
    density_rank_in_province = query_results['Density_rank_in_province']  # 本地区密度在全省排名
    stats_of_region = query_results['Stats_of_Region']  # 分县市统计
    day_target_area = stats_of_region[u'总计'][4]  # 平均雷暴日
    day_max_target = stats_of_region[u'总计'][5]  # 最大雷暴日
    day_min_target = stats_of_region[u'总计'][6]  # 最小雷暴日

    sum_max_county_name = query_results['Sum_max_county_name']  # 闪电次数最多的县名
    sum_max_in_region = query_results['Sum_max_in_region']  # 闪电次数最多的县次数
    sum_min_county_name = query_results['Sum_min_county_name']  # 闪电次数最少的县名
    sum_min_in_region = query_results['Sum_min_in_region']  # 闪电次数最少的县次数
    density_max_county_name = query_results['Density_max_county_name']  # 闪电密度最多的县名
    density_max_in_region = query_results['Density_max_in_region']  # 闪电密度最多的县密度
    density_min_county_name = query_results['Density_min_county_name']  # 闪电密度最少的县名
    density_min_in_region = query_results['Density_min_in_region']  # 闪电密度最少的县密度
    max_region_percent = query_results['Max_region_percent']  # 地闪次数最大县所占比例
    min_region_percent = query_results['Min_region_percent']  # 地闪次数最小县所占比例

    peak_month_negative_intensity = query_results['Peak_month_negative_intensity']  # 正闪峰值月份
    peak_month_positive_intensity = query_results['Peak_month_positive_intensity']  # 负闪峰值月份
    max_month_region = query_results['Max_month_region']  # 地闪次数最多的月份
    max_three_months = query_results['Max_three_months']  # 地闪次数最多的三个月
    max_months_percent = query_results['Max_months_percent']  # 地闪次数最多三个月所占比例
    months_zero = query_results['Months_zero']  # 没有检测到地闪的月份
    stats_of_Month = query_results['Stats_of_Month']  # 分月统计，负闪次数、负闪强度、正闪次数、正闪强度、总次数
    first_date = query_results['First_date']  # 雷暴初日

    stats_of_province = query_results[u'Stats_of_Province']

    # ******页眉页脚处理*******
    # 设置首页、奇偶页的页眉页脚全部一样
    doc.PageSetup.DifferentFirstPageHeaderFooter = False
    doc.PageSetup.OddAndEvenPagesHeaderFooter = False
    # 整体替换页眉
    header = doc.Sections(2).Headers(1).Range
    header.Find.ClearFormatting()
    header.Find.Replacement.ClearFormatting()
    header.Find.Execute(u'2015年', False, False, False, False, False, True, 1, False, year, 2)
    header.Find.Execute(u'绍兴市', False, False, False, False, False, True, 1, False, target_area, 2)
    # 整体替换页脚
    footer = doc.Sections(2).Footers(1).Range
    footer.Find.ClearFormatting()
    footer.Find.Replacement.ClearFormatting()
    footer.Find.Execute(u'2015年', False, False, False, False, False, True, 1, False, year, 2)
    footer.Find.Execute(u'绍兴市', False, False, False, False, False, True, 1, False, target_area, 2)


    word.Selection.Find.ClearFormatting()
    word.Selection.Find.Replacement.ClearFormatting()
    word.Selection.Find.Execute(u'2016年',False, False, False, False, False, True, 1, True, year,2)

    # ******替换图片*********
    word.Selection.Find.Execute(FindText=u'图1-1 地闪密度空间分布图', Wrap=1)
    word.Selection.MoveLeft(Count=3)
    doc.InlineShapes(2).Delete()
    densityPic = ''.join([cwd, u'/temp/', province, u'/', year, u'/', target_area, u'.gdb/',
                          year, target_area, u'闪电密度空间分布.png'])
    word.Selection.InlineShapes.AddPicture(densityPic)

    word.Selection.Find.Execute(FindText=u'图1-2 地闪雷暴日空间分布图', Wrap=1)
    word.Selection.MoveLeft(Count=3)
    doc.InlineShapes(3).Delete()
    dayPic = ''.join([cwd, u'/temp/', province, u'/', year, u'/', target_area, u'.gdb/',
                      year, target_area, u'地闪雷暴日空间分布.png'])
    word.Selection.InlineShapes.AddPicture(dayPic)

    # 在word中定位
    word.Selection.Find.Execute(FindText=u'图1-3 地闪分月统计', Wrap=1)
    word.Selection.MoveLeft(Count=3)
    # 删除原有图表
    doc.InlineShapes(4).Delete()
    # 复制Excel中图表
    sheet = workbook.Worksheets(u'分月统计')
    sheet.ChartObjects(1).Chart.ChartArea.Copy()
    # 在word中粘贴
    word.Selection.PasteAndFormat(1)  # wdChartLinked

    word.Selection.Find.Execute(FindText=u'图1-4 地闪分时段统计', Wrap=1)
    word.Selection.MoveLeft(Count=3)
    doc.InlineShapes(5).Delete()
    sheet = workbook.Worksheets(u'分时段统计')
    sheet.ChartObjects(1).Chart.ChartArea.Copy()
    word.Selection.PasteAndFormat(1)  # wdChartLinked

    word.Selection.Find.Execute(FindText=u'图1-5 负地闪强度分布', Wrap=1)
    word.Selection.MoveLeft(Count=3)
    doc.InlineShapes(6).Delete()
    sheet = workbook.Worksheets(u'强度分布统计')
    sheet.ChartObjects(1).Chart.ChartArea.Copy()
    word.Selection.PasteAndFormat(1)  # wdChartLinked

    word.Selection.Find.Execute(FindText=u'图 1-6 正地闪强度分布', Wrap=1)
    word.Selection.MoveLeft(Count=3)
    doc.InlineShapes(7).Delete()
    sheet.ChartObjects(2).Chart.ChartArea.Copy()
    word.Selection.PasteAndFormat(1)  # wdChartLinked

    # ************处理表格**************
    if province == u'浙江':
        regions = [u'杭州', u'宁波', u'湖州', u'嘉兴', u'绍兴', u'金华', u'台州', u'温州', u'衢州', u'丽水', u'舟山', u'总计']
        countries = [u'越城区', u'柯桥区', u'上虞区', u'诸暨市', u'嵊州市', u'新昌县', u'总计']
        n_regions = len(regions)
        n_countries = len(countries)

        table_region = doc.Tables(1)
        for i in range(n_countries):
            table_region.Cell(2, i + 2).Range.Text = str(stats_of_region[countries[i]][0])  # 地闪次数
            table_region.Cell(3, i + 2).Range.Text = str(stats_of_region[countries[i]][1])  # 平均地闪密度
            table_region.Cell(4, i + 2).Range.Text = str(stats_of_region[countries[i]][4])  # 平均雷暴日

        table_province = doc.Tables(2)
        for i in range(n_regions):
            for j in range(2):
                table_province.Cell(j + 2, i + 2).Range.Text = str(stats_of_province[regions[i]][j])

    elif province == u'河南':
        regions = [u'郑州', u'开封', u'洛阳', u'平顶山', u'安阳', u'鹤壁', u'新乡', u'焦作', u'濮阳',
                   u'许昌', u'漯河', u'三门峡', u'商丘', u'周口', u'驻马店', u'南阳', u'信阳', u'济源', u'总计']
        countries = [u'市区', u'新乡县', u'辉县市', u'卫辉市', u'获嘉县', u'原阳县', u'延津县', u'封丘县', u'长垣县', u'总计']
        n_regions = len(regions)
        n_countries = len(countries)

        table_region = doc.Tables(1)
        for i in range(n_countries):
            table_region.Cell(2, i + 2).Range.Text = str(stats_of_region[countries[i]][0])  # 地闪次数
            table_region.Cell(3, i + 2).Range.Text = str(stats_of_region[countries[i]][1])  # 平均地闪密度
            table_region.Cell(4, i + 2).Range.Text = str(stats_of_region[countries[i]][4])  # 平均雷暴日

        table_province = doc.Tables(2)
        for i in range(10):
            for j in range(2):
                table_province.Cell(j + 2, i + 2).Range.Text = str(stats_of_province[regions[i]][j])
        table_province_xu = doc.Tables(3)
        for i in range(9):
            for j in range(2):
                table_province_xu.Cell(j + 2, i + 2).Range.Text = str(stats_of_province[regions[i + 10]][j])

    # ****************段落文字处理*******************
    # 处理与去年的对比情况
    query_results_path = ''.join([cwd, u"/temp/", province, '/', year, '/', target_area, '.gdb/', 'query_results.pkl'])
    last_year_query_results_path = query_results_path.replace(year, str(int(year[:-1]) - 1) + u'年')
    if os.path.exists(last_year_query_results_path):
        f = open(last_year_query_results_path, 'rb')
        sum_target_area_last_year = (pickle.load(f))['Sum_target_area']
        f.close()

        rate = (sum_target_area - sum_target_area_last_year) / float(sum_target_area_last_year)

        if 0.05 <= rate < 0.1:
            compare = u'略有增多'
        elif 0.1 <= rate < 0.3:
            compare = u'有所增多'
        elif 0.3 <= rate < 0.9:
            compare = u'增幅较大'
        elif 0.9 <= rate:
            compare = u'大幅增多'

        elif -0.1 < rate <= -0.05:
            compare = u'略有减少'
        elif -0.3 < rate <= -0.1:
            compare = u'有所减少'
        elif -0.9 < rate <= -0.3:
            compare = u'减幅较大'
        elif rate <= -0.9:
            compare = u'大幅减少'
        else:
            compare = u'基本持平'

        compare_with_last_year = u'与上年的地闪%d次相比，%s。' % (sum_target_area_last_year, compare)
    else:
        compare_with_last_year = ''

    if density_target_area > density_province:
        compare_with_province = u'高于'
    else:
        compare_with_province = u'低于'

    p1 = u'%s我市共发生地闪%d次，平均地闪密度%.2f次/km²，平均雷暴日%d天（见表1-1）。%s\
从时间分布来看，地闪主要集中在%d、%d、%d月，三个月地闪占全年总地闪次数的%.2f%%。从空间分布来看，%s发生地闪次数最多，%s最少。\
全市地闪平均密度%s全省平均的%.2f次/km²，在全省各市中%s闪次数排第%d位，地闪平均密度排第%d位（见表1-2）。' % (
        year, sum_target_area, density_target_area, day_target_area, compare_with_last_year,
        max_three_months[0], max_three_months[1], max_three_months[2], max_months_percent,
        sum_max_county_name, sum_min_county_name, compare_with_province, density_province,
        target_area, sum_rank_in_province, density_rank_in_province)

    p2 = u'据不完全统计，2016年全市因雷电引发的灾害共148起，无人员伤亡事故。\
造成直接经济损失达7788.04万元，间接经济损失677.42万元。'

    p3 = u'从地区统计来看，地区分布相对不均，%s地闪次数最多，共%d次，%s最少，只有%d次，\
两者分别占全市总地闪数的%.2f%%和%.2f%%。从平均密度统计来看，%s密度最高，为%.2f次/km²，\
%s最低，为%.2f次/km²（见表1-1）。' % (sum_max_county_name, sum_max_in_region, sum_min_county_name, sum_min_in_region,
                            max_region_percent, min_region_percent,
                            density_max_county_name, density_max_in_region,
                            density_min_county_name, density_min_in_region)

    p4 = u'从地闪密度空间分布图上（见图1-1）可以看出，诸暨西北部、嵊州和诸暨交界区域地闪密度较高，\
最高超过5次/km²。新昌东部有部分地区，地闪密度超过3次/km²，全市大部分地区地闪密度小于2次/km²。'

    p5 = u'现行国家标准所引用的雷暴日指人工观测（测站周围约15km半径域面）有雷暴天数的多年平均。\
根据我省闪电定位监测资料推算（以15km为间隔，分别统计各点15km半径范围内的雷暴日，再插值推算），\
%s全市地闪雷暴日平%d天，最低为%d天，最高%d天。空间分布上来看，北部平原地区雷暴日较少，\
西南大部和东南部分区域雷暴天数较多（见图1-2）。' % (year, day_target_area, day_min_target, day_max_target)

    if len(months_zero) == 0:
        months_zero_description = u''
    elif len(months_zero) == 1:
        months_zero_description = u'%d月未监测到地闪，' % months_zero[0]
    elif len(months_zero) == 2:
        months_zero_description = u'%d月和%d月未监测到地闪，' % (months_zero[0], months_zero[1])
    else:
        s = u'月、'.join(map(lambda d: str(d), months_zero[:len(months_zero) - 1]))
        months_zero_description = u''.join([s, u'月和%d月都未监测到地闪，' % months_zero[-1]])

    p6 = u'%s%s雷电初日为%s。从分月统计来看，地闪次数随月份呈现近似正态分布特征，%s\
地闪次数峰值出现在%d月，%d、%d、%d月是雷暴高发的月份，三个月地闪次数占总数的%.2f%%。\
正、负地闪平均强度的峰值分别在%d月和%d月，其他月份波动平缓(见图1-3) 。' % (year, target_area,
                                             first_date, months_zero_description,
                                             max_month_region, max_three_months[0], max_three_months[1],
                                             max_three_months[2],
                                             max_months_percent,
                                             peak_month_positive_intensity, peak_month_negative_intensity)

    p7 = u'从分时段统计来看，地闪次数峰值出现在第18个时段（17:00-18:00），地闪主要集中在午后两点到晚上九点，\
七个时段内的地闪次数占总数的d%。地闪平均强度随时间呈波状起伏特征，但总体波动不大。\
正闪平均强度峰值在第7个时段（7:00-8:00），负闪平均强度峰值在第11个时段（11:00-12:00）(见图1-4)。'

    p8 = u'由正、负地闪强度分布图可见，地闪次数随地闪强度呈近似正态分布特征。正地闪主要集中在5-60kA内（见图1-5），\
该区间内正地闪次数约占总地闪的87.20%，负地闪主要分布在5-60kA内（见图1-6），\
该区间内负地闪次数约占总负地闪的91.46%。'

    paragraphs = [p1, p2, p3, p4, p5, p6, p7, p8]
    for i in range(8):
        word.Selection.Find.Execute(FindText=u'#段落%d#' % (i + 1), Wrap=1)
        word.Selection.TypeText(Text=paragraphs[i])


def mainProcess(year, province, target_area, cwd):
    # 打开Excel应用程序
    excel = DispatchEx('Excel.Application')
    excel.Visible = False
    # 打开文件，即Excel工作薄
    charts_origin = ''.join([cwd, u'/data/', u'公报图表模板.xlsx'])
    charts = ''.join([cwd, u'/temp/', province, '/', year, '/', target_area, '.gdb/',
                      year, target_area, u'公报统计图表.xlsx'])

    shutil.copy2(charts_origin, charts)
    workbook = excel.Workbooks.Open(charts)

    # 打开word文档
    EnsureDispatch('Word.Application')
    word = DispatchEx('Word.Application')
    word.Visible = False

    doc_origin = ''.join([cwd, u'/data/', u'公报文档模板_%s.docx' % target_area])
    doc = ''.join([cwd, u'/temp/', province, '/', year, '/', target_area, '.gdb/',
                   year, target_area, u'公报文档.docx'])

    # if not os.path.exists(doc):
    shutil.copy2(doc_origin, doc)
    doc = word.Documents.Open(doc)

    try:
        query_results = sqlQuery(workbook, year, province, target_area, cwd)
        docProcess(word, doc, workbook, query_results, year, province, target_area, cwd)
    finally:
        doc.Save()
        doc.Close()
        word.Quit()
        workbook.Save()  # 保存EXCEL工作薄
        workbook.Close()  # 关闭工作薄文件
        excel.Quit()  # 关闭EXCEL应用程序


if __name__ == "__main__":

    (year, province, target_area,cwd) = sys.argv[1:]
    mainProcess(year, province, target_area,cwd)

    # year = u"2014年"
    # province = u'浙江'
    # target_area = u"绍兴市"
    #
    # cwd = os.getcwd()
    # start = time.clock()
    # # ***********************测试程序*********************************
    # mainProcess(year, province, target_area, cwd)
    # # docProcess(year, province, target_area, cwd)
    # # ***********************测试程序*********************************
    # end = time.clock()
    # elapsed = end - start
    # print("Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000))
