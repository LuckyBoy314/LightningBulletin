# -*- coding: gbk -*-

import pyodbc
import os
import time
from win32com.client import DispatchEx
import pickle
import sys

def dict2list(dic: dict):
    ''' 将字典转化为列表 '''
    keys = dic.keys()
    vals = dic.values()
    lst = [(key, val) for key, val in zip(keys, vals)]
    return lst


def sqlQuery(year, province, target_area, cwd):
    query_results = {}

    # 链接数据库
    database = ''.join([cwd, u"/temp/", province, '/', year, '/SQL.mdb;'])

    db = pyodbc.connect(''.join(['DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};', 'DBQ=', database]))
    cursor = db.cursor()
    data_table = ''.join(['data', year])  # 查询表

    # 打开Excel应用程序
    excel = DispatchEx('Excel.Application')
    excel.Visible = False
    # 打开文件，即Excel工作薄
    workbook = excel.Workbooks.Open(''.join([cwd, u'/data/公报图表模板.xlsx']))

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
        Density_target_area = round(Sum_target_area / province_area[target_area],2) # 本地区地闪密度

        # 计算全省各地区地闪密度和全省平均密度
        density_province_dict = {}  # 全省各地区密度
        Density_province = 0
        for key in sum_province_dict:
            density_province_dict[key] = round(sum_province_dict[key] / province_area[key],2)
            Density_province += density_province_dict[key]
        Density_province = round(Density_province/len(province_area),2)  # 全省平均地闪密度
        # 密度从大到小进行排序
        density_province_sorted = sorted(dict2list(density_province_dict), key=lambda d: d[1], reverse=True)

        # 计算本地区地闪密度排名
        rank = 0
        for item in density_province_sorted:
            rank += 1
            if item[0] == target_area:
                Density_rank_in_province = rank  # 本地区地闪密度在全省的排名
                break

        #省分地区的电闪次数和密度
        Stats_of_Province = {}
        sum_of_province = 0
        for key in sum_province_dict:
            sum_of_province += sum_province_dict[key]
            Stats_of_Province[key[:2]] = (sum_province_dict[key], density_province_dict[key])
        Stats_of_Province[u'总计'] = (sum_of_province, Density_province)

        # print('本地区地闪总数：', Sum_target_area)
        # print('本地区地闪总数在全省排名：', Sum_rank_in_province)
        # print('全省密度：', Density_province)
        # print('本地区密度：', Density_target_area)
        # print('本地区密度在全省排名：', Density_rank_in_province)
        # print('省分地区统计：', Stats_of_Province)

        query_results['Sum_target_area'] = Sum_target_area
        query_results['Sum_rank_in_province'] = Sum_rank_in_province
        query_results['Density_province'] = Density_province
        query_results['Density_target_area'] = Density_target_area
        query_results['Density_rank_in_province'] = Density_rank_in_province
        query_results['Stats_of_Province'] = Stats_of_Province

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
        Max_region_percent = round(Sum_max_in_region / float(Sum_target_area) * 100,2)
        Min_region_percent = round(Sum_min_in_region / float(Sum_target_area) * 100,2)

        # 计算本地区各县市地闪密度
        density_region_dict = {}
        for key in sum_region_dict:
            density_region_dict[key] = round(sum_region_dict[key] / region_area[key],2)
        # 密度从大到小进行排序
        density_region_sorted = sorted(dict2list(density_region_dict), key=lambda d: d[1], reverse=True)
        # 最大、最小密度
        Density_max_county_name = density_region_sorted[0][0]  # 闪电密度最高的县名
        Density_max_in_region = density_region_sorted[0][1]  # 闪电密度最高的县密度
        Density_min_county_name = density_region_sorted[num_region - 1][0]  # 闪电密度最低的县名
        Density_min_in_region = density_region_sorted[num_region - 1][1]  # 闪电密度最低的县密度

        Stats_of_Region = {}#地闪次数、平均地闪密度、最大地闪密度、最小地闪密度、平均雷暴日、最大雷暴日、最小雷暴日
        for key in sum_region_dict:
            Stats_of_Region[key] = (sum_region_dict[key], density_region_dict[key],
                                    stat_density[key[:2]][0],stat_density[key[:2]][1],
                                    stat_day[key[:2]][0],stat_day[key[:2]][1],stat_day[key[:2]][2])
        max_density_region = max([stat_density[each][0] for each in stat_density])
        min_density_region = min([stat_density[each][1] for each in stat_density])
        Day_target_area = int(sum([stat_day[each][0] for each in stat_day])/num_region)
        max_day_region = max([stat_day[each][1] for each in stat_day])
        min_day_region = min([stat_day[each][2] for each in stat_day])

        Stats_of_Region[u'总计'] = (Sum_target_area,Density_target_area,max_density_region,min_density_region,
                                  Day_target_area,max_day_region,min_day_region)

        # print('闪电次数最多的县名:', Sum_max_county_name)
        # print('闪电次数最多的县次数:', Sum_max_in_region)
        # print('闪电次数最少的县名:', Sum_min_county_name)
        # print('闪电次数最少的县次数:', Sum_min_in_region)
        #
        # print('闪电密度最多的县名:', Density_max_county_name)
        # print('闪电密度最多的县密度:', Density_max_in_region)
        # print('闪电密度最少的县名:', Density_min_county_name)
        # print('闪电密度最少的县密度:', Density_min_in_region)
        # print('地闪次数最大县所占比例：', Max_region_percent)
        # print('地闪次数最小县所占比例：', Min_region_percent)
        # print('分县市统计：', Stats_of_Region)

        query_results['Sum_max_county_name'] = Sum_max_county_name
        query_results['Sum_max_in_region'] = Sum_max_in_region
        query_results['Sum_min_county_name'] = Sum_min_county_name
        query_results['Sum_min_in_region'] = Sum_min_in_region
        query_results['Density_max_county_name'] = Density_max_county_name
        query_results['Density_max_in_region'] = Density_max_in_region
        query_results['Density_min_county_name'] = Density_min_county_name
        query_results['Density_min_in_region'] = Density_min_in_region
        query_results['Max_region_percent'] = Max_region_percent
        query_results['Min_region_percent'] = Min_region_percent
        query_results['Stats_of_Region'] = Stats_of_Region

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
            Stats_of_Month[month] = [row[0],negative_intensity_dict[month]]

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
        Max_months_percent = round(100 * (sum_month_sorted[0][1] + sum_month_sorted[1][1] + sum_month_sorted[2][1]) / float(
            Sum_target_area),2)
        # 没有检测到地闪的月份
        Months_zero = [i[0] for i in sum_month_sorted if i[1] == 0]
        Months_zero.sort()

        # print('正闪峰值月份:', Peak_month_negative_intensity)
        # print('负闪峰值月份:', Peak_month_positive_intensity)
        # print('地闪次数最多的月份:', Max_month_region)
        # print('地闪次数最多的三个月:', Max_three_months)
        # print('地闪次数最多三个月所占比例:', Max_months_percent)
        # print('没有检测到地闪的月份:', Months_zero)

        query_results['Peak_month_negative_intensity'] = Peak_month_negative_intensity
        query_results['Peak_month_positive_intensity'] = Peak_month_positive_intensity
        query_results['Max_month_region'] = Max_month_region
        query_results['Max_three_months'] = Max_three_months
        query_results['Max_months_percent'] = Max_months_percent
        query_results['Months_zero'] = Months_zero
        query_results['Stats_of_Month'] = Stats_of_Month #负闪次数、负闪强度、正闪次数、正闪强度、总次数

        # ****查询雷暴初日********
        sql = """SELECT TOP 1 Date_
        From %s
        Where Region = \'%s\'
        Order By Date_, OBJECTID
        """ % (data_table, target_area)

        for row in cursor.execute(sql):
            First_date = row[0].strftime('%mM%dD').replace('M', '月').replace('D', '日')

        # print('雷暴初日:', First_date)
        query_results['First_date'] = First_date

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
            hour = i-2
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
            hour = i-2
            sheet.Cells(i, 3).Value = row[0]  # 正闪次数
            sheet.Cells(i, 6).Value = positive_intensity_hours = row[1] if row[1] is not None else 0  # 正闪强度
            sum_hour_dict[hour] += row[0]
            Stats_of_Hour[hour].append(row[0])
            Stats_of_Hour[hour].append(positive_intensity_hours)
            Stats_of_Hour[hour].append(sum_hour_dict[hour])

        query_results['Stats_of_Hour'] = Stats_of_Hour

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

        sheet = workbook.Worksheets(u'强度分布统计')
        i = 1  # 行号
        for row in cursor.execute(sql):
            i += 1
            sheet.Cells(i, 3).Value = row[0]  # 负闪次数

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

        f = open(os.path.join(workspath, 'query_results.pkl'), 'wb')
        pickle.dump(query_results, f, 2)#不能采用最高的协议，否则在python2.7.8中无法加载进来
        f.close()


    finally:
        db.close()  # 关闭数据连接
        workbook.Save()  # 保存EXCEL工作薄
        workbook.Close()  # 关闭工作薄文件
        excel.Quit()  # 关闭EXCEL应用程序


if __name__ == "__main__":

    (year, province, target_area,cwd) = sys.argv[1:]
    sqlQuery(year, province, target_area,cwd)

    # year = u"2015年"
    # province = u'浙江'
    # target_area = u"绍兴市"
    #
    # cwd = os.getcwd()
    # start = time.clock()
    # # ***********************测试程序*********************************"
    # sqlQuery(year, province, target_area, cwd)
    # # ***********************测试程序*********************************"
    # end = time.clock()
    # elapsed = end - start
    # print("Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000))
