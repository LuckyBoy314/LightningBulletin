# -*- coding: gbk -*-

import pyodbc
import os
import time
from win32com.client import DispatchEx
import pickle
import sys

def dict2list(dic: dict):
    ''' ���ֵ�ת��Ϊ�б� '''
    keys = dic.keys()
    vals = dic.values()
    lst = [(key, val) for key, val in zip(keys, vals)]
    return lst


def sqlQuery(year, province, target_area, cwd):
    query_results = {}

    # �������ݿ�
    database = ''.join([cwd, u"/temp/", province, '/', year, '/SQL.mdb;'])

    db = pyodbc.connect(''.join(['DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};', 'DBQ=', database]))
    cursor = db.cursor()
    data_table = ''.join(['data', year])  # ��ѯ��

    # ��ExcelӦ�ó���
    excel = DispatchEx('Excel.Application')
    excel.Visible = False
    # ���ļ�����Excel������
    workbook = excel.Workbooks.Open(''.join([cwd, u'/data/����ͼ��ģ��.xlsx']))

    # ��ȡʡ�������ؼ������
    workspath = ''.join([cwd, u"/temp/", province, '/', year, '/', 'GDB.gdb'])
    f = open(os.path.join(workspath, 'province_area.pkl'), 'rb')
    province_area = pickle.load(f)  # ����{'������': XXX, '������': XXX}
    f.close()

    workspath = ''.join([cwd, u"/temp/", province, '/', year, '/', target_area, '.gdb'])
    # ��ȡ�ؼ����������������
    f = open(os.path.join(workspath, 'region_area.pkl'), 'rb')
    region_area = pickle.load(f)  # ����{'������'��XXX, '�ӽ���':XXX}
    f.close()

    # ��ȡĿ��ؼ�������ױ��պ������ܶȵ�ͳ����Ϣ
    f = open(os.path.join(workspath, 'stats.pkl'), 'rb')
    stat_density = pickle.load(f)  # ����{'����':(���ֵ����Сֵ)}
    stat_day = pickle.load(f)  # ����{'����':(ƽ��ֵ�����ֵ����Сֵ)}
    f.close()

    try:
        # ************�ֵ���ͳ��**********
        sql = """
        SELECT count(*) AS num, Region
        FROM %s
        WHERE Province= \'%s\'
        GROUP BY Region
        ORDER BY count(*) DESC
        """ % (data_table, province + 'ʡ')

        # ����SQL��ѯ�����˳���¼����������������ȫʡ������
        sum_province_dict = {}  # ȫʡ��������������
        rank = 0
        for row in cursor.execute(sql):
            sum_province_dict[row[1]] = row[0]  # �� Region��num�����ֵ䣬�������渳ֵ
            rank += 1
            # print type(row[1])
            if row[1] == target_area:
                Sum_rank_in_province = rank  # ����������������ȫʡ������

        Sum_target_area = sum_province_dict[target_area]  # ��������������
        Density_target_area = round(Sum_target_area / province_area[target_area],2) # �����������ܶ�

        # ����ȫʡ�����������ܶȺ�ȫʡƽ���ܶ�
        density_province_dict = {}  # ȫʡ�������ܶ�
        Density_province = 0
        for key in sum_province_dict:
            density_province_dict[key] = round(sum_province_dict[key] / province_area[key],2)
            Density_province += density_province_dict[key]
        Density_province = round(Density_province/len(province_area),2)  # ȫʡƽ�������ܶ�
        # �ܶȴӴ�С��������
        density_province_sorted = sorted(dict2list(density_province_dict), key=lambda d: d[1], reverse=True)

        # ���㱾���������ܶ�����
        rank = 0
        for item in density_province_sorted:
            rank += 1
            if item[0] == target_area:
                Density_rank_in_province = rank  # �����������ܶ���ȫʡ������
                break

        #ʡ�ֵ����ĵ����������ܶ�
        Stats_of_Province = {}
        sum_of_province = 0
        for key in sum_province_dict:
            sum_of_province += sum_province_dict[key]
            Stats_of_Province[key[:2]] = (sum_province_dict[key], density_province_dict[key])
        Stats_of_Province[u'�ܼ�'] = (sum_of_province, Density_province)

        # print('����������������', Sum_target_area)
        # print('����������������ȫʡ������', Sum_rank_in_province)
        # print('ȫʡ�ܶȣ�', Density_province)
        # print('�������ܶȣ�', Density_target_area)
        # print('�������ܶ���ȫʡ������', Density_rank_in_province)
        # print('ʡ�ֵ���ͳ�ƣ�', Stats_of_Province)

        query_results['Sum_target_area'] = Sum_target_area
        query_results['Sum_rank_in_province'] = Sum_rank_in_province
        query_results['Density_province'] = Density_province
        query_results['Density_target_area'] = Density_target_area
        query_results['Density_rank_in_province'] = Density_rank_in_province
        query_results['Stats_of_Province'] = Stats_of_Province

        # ********* ����ͳ��***********
        sql = """
        SELECT count(*) AS num, County
        FROM %s
        WHERE Region=\'%s\'
        GROUP BY County
        ORDER BY count(*) DESC
        """ % (data_table, target_area)

        # ����SQL��ѯ�����˳���¼����������������Сֵ
        sum_region_dict = {}  # �������������ص�������
        rank = 0
        num_region = len(region_area)
        for row in cursor.execute(sql):
            sum_region_dict[row[1]] = row[0]  # �� Region��num�����ֵ䣬�������渳ֵ
            rank += 1
            if rank == 1:
                Sum_max_county_name = row[1]  # ���������������
                Sum_max_in_region = row[0]  # �Լ�����
            elif rank == num_region:
                Sum_min_county_name = row[1]  # ����������ٵ�����
                Sum_min_in_region = row[0]  # �Լ�����

        # ����������������С���ص�ռĿ�����������ı���
        Max_region_percent = round(Sum_max_in_region / float(Sum_target_area) * 100,2)
        Min_region_percent = round(Sum_min_in_region / float(Sum_target_area) * 100,2)

        # ���㱾���������е����ܶ�
        density_region_dict = {}
        for key in sum_region_dict:
            density_region_dict[key] = round(sum_region_dict[key] / region_area[key],2)
        # �ܶȴӴ�С��������
        density_region_sorted = sorted(dict2list(density_region_dict), key=lambda d: d[1], reverse=True)
        # �����С�ܶ�
        Density_max_county_name = density_region_sorted[0][0]  # �����ܶ���ߵ�����
        Density_max_in_region = density_region_sorted[0][1]  # �����ܶ���ߵ����ܶ�
        Density_min_county_name = density_region_sorted[num_region - 1][0]  # �����ܶ���͵�����
        Density_min_in_region = density_region_sorted[num_region - 1][1]  # �����ܶ���͵����ܶ�

        Stats_of_Region = {}#����������ƽ�������ܶȡ��������ܶȡ���С�����ܶȡ�ƽ���ױ��ա�����ױ��ա���С�ױ���
        for key in sum_region_dict:
            Stats_of_Region[key] = (sum_region_dict[key], density_region_dict[key],
                                    stat_density[key[:2]][0],stat_density[key[:2]][1],
                                    stat_day[key[:2]][0],stat_day[key[:2]][1],stat_day[key[:2]][2])
        max_density_region = max([stat_density[each][0] for each in stat_density])
        min_density_region = min([stat_density[each][1] for each in stat_density])
        Day_target_area = int(sum([stat_day[each][0] for each in stat_day])/num_region)
        max_day_region = max([stat_day[each][1] for each in stat_day])
        min_day_region = min([stat_day[each][2] for each in stat_day])

        Stats_of_Region[u'�ܼ�'] = (Sum_target_area,Density_target_area,max_density_region,min_density_region,
                                  Day_target_area,max_day_region,min_day_region)

        # print('���������������:', Sum_max_county_name)
        # print('������������ش���:', Sum_max_in_region)
        # print('����������ٵ�����:', Sum_min_county_name)
        # print('����������ٵ��ش���:', Sum_min_in_region)
        #
        # print('�����ܶ���������:', Density_max_county_name)
        # print('�����ܶ��������ܶ�:', Density_max_in_region)
        # print('�����ܶ����ٵ�����:', Density_min_county_name)
        # print('�����ܶ����ٵ����ܶ�:', Density_min_in_region)
        # print('���������������ռ������', Max_region_percent)
        # print('����������С����ռ������', Min_region_percent)
        # print('������ͳ�ƣ�', Stats_of_Region)

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

        # todo SQL��ѯ�д��Ż�
        # ************����ͳ�� �µ�����������ƽ��ǿ��(����)**************
        sql = """
        SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
        UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
        ORDER BY �·�
        """.replace('TARGET_AREA', target_area).replace('YEAR', year[0:4]).replace('QUERY_TABLE', data_table)

        sheet = workbook.Worksheets(u'����ͳ��')
        Stats_of_Month = {}
        i = 1  # �к�
        month = 0
        sum_month_dict = {}  # ��¼���µ�������
        negative_intensity_dict = {}  # ��¼���¸�����������
        for row in cursor.execute(sql):
            i += 1
            month += 1
            sheet.Cells(i, 2).Value = row[0]  # ��������
            sheet.Cells(i, 5).Value = negative_intensity_dict[month] = row[1] if row[1] is not None else 0  # ����ǿ��
            sum_month_dict[month] = row[0]
            Stats_of_Month[month] = [row[0],negative_intensity_dict[month]]

        # ����ǿ�ȷ�ֵ�����·�
        negative_intensity_sorted = sorted(dict2list(negative_intensity_dict), key=lambda d: d[1], reverse=True)
        Peak_month_negative_intensity = negative_intensity_sorted[0][0]

        # ************����ͳ�� �µ�����������ƽ��ǿ��(����)**************
        sql = """
        SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS �·�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
        ORDER BY �·�
        """.replace('TARGET_AREA', target_area).replace('YEAR', year[0:4]).replace('QUERY_TABLE', data_table)

        i = 1  # �к�
        month = 0
        positive_intensity_dict = {}  # ��¼����ǿ��
        for row in cursor.execute(sql):
            i += 1
            month += 1
            sheet.Cells(i, 3).Value = row[0]  # ��������
            sheet.Cells(i, 6).Value = positive_intensity_dict[month] = row[1] if row[1] is not None else 0  # ����ǿ��
            sum_month_dict[month] += row[0]
            Stats_of_Month[month].append(row[0])
            Stats_of_Month[month].append(positive_intensity_dict[month])
            Stats_of_Month[month].append(sum_month_dict[month])

        # ������ֵ�·�
        positive_intensity_sorted = sorted(dict2list(positive_intensity_dict), key=lambda d: d[1], reverse=True)
        Peak_month_positive_intensity = positive_intensity_sorted[0][0]

        sum_month_sorted = sorted(dict2list(sum_month_dict), key=lambda d: d[1], reverse=True)
        # �������������·�
        Max_month_region = sum_month_sorted[0][0]
        # ������������������
        Max_three_months = [sum_month_sorted[0][0], sum_month_sorted[1][0], sum_month_sorted[2][0]]
        Max_three_months.sort()
        # �������������������ռ����
        Max_months_percent = round(100 * (sum_month_sorted[0][1] + sum_month_sorted[1][1] + sum_month_sorted[2][1]) / float(
            Sum_target_area),2)
        # û�м�⵽�������·�
        Months_zero = [i[0] for i in sum_month_sorted if i[1] == 0]
        Months_zero.sort()

        # print('������ֵ�·�:', Peak_month_negative_intensity)
        # print('������ֵ�·�:', Peak_month_positive_intensity)
        # print('�������������·�:', Max_month_region)
        # print('������������������:', Max_three_months)
        # print('�������������������ռ����:', Max_months_percent)
        # print('û�м�⵽�������·�:', Months_zero)

        query_results['Peak_month_negative_intensity'] = Peak_month_negative_intensity
        query_results['Peak_month_positive_intensity'] = Peak_month_positive_intensity
        query_results['Max_month_region'] = Max_month_region
        query_results['Max_three_months'] = Max_three_months
        query_results['Max_months_percent'] = Max_months_percent
        query_results['Months_zero'] = Months_zero
        query_results['Stats_of_Month'] = Stats_of_Month #��������������ǿ�ȡ���������������ǿ�ȡ��ܴ���

        # ****��ѯ�ױ�����********
        sql = """SELECT TOP 1 Date_
        From %s
        Where Region = \'%s\'
        Order By Date_, OBJECTID
        """ % (data_table, target_area)

        for row in cursor.execute(sql):
            First_date = row[0].strftime('%mM%dD').replace('M', '��').replace('D', '��')

        # print('�ױ�����:', First_date)
        query_results['First_date'] = First_date

        # ************��ʱ��ͳ�� ʱ�ε���������ʱ��ƽ��ǿ��(����)**************
        sql = """
        SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,0 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=0
        UNION SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=1
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=2
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=3
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=4
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=5
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=6
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=7
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=8
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=9
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=10
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=11
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=12
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,13 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=13
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,14 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=14
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,15 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=15
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,16 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=16
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,17 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=17
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,18 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=18
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,19 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=19
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,20 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=20
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,21 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=21
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,22 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=22
        union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,23 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Val(Time_)=23
        ORDER BY ʱ��
        """.replace('TARGET_AREA', target_area).replace('QUERY_TABLE', data_table)

        sheet = workbook.Worksheets(u'��ʱ��ͳ��')
        sum_hour_dict = {}  # ��¼��ʱ�ε�������
        Stats_of_Hour = {}
        i = 1  # �к�
        for row in cursor.execute(sql):
            i += 1
            hour = i-2
            sheet.Cells(i, 2).Value = row[0]  # ��������
            sheet.Cells(i, 5).Value = negative_intensity_hours = row[1] if row[1] is not None else 0  # ����ǿ��
            sum_hour_dict[hour] = row[0]
            Stats_of_Hour[hour] = [row[0], negative_intensity_hours]



        # ************��ʱ��ͳ�� ʱ�ε���������ʱ��ƽ��ǿ��(����)**************
        sql = """
        SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,0 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=0
        UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=1
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=2
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=3
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=4
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=5
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=6
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=7
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=8
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=9
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=10
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=11
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=12
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,13 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=13
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,14 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=14
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,15 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=15
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,16 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=16
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,17 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=17
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,18 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=18
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,19 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=19
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,20 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=20
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,21 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=21
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,22 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=22
        union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,23 AS ʱ��
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Val(Time_)=23
        ORDER BY ʱ��
        """.replace('TARGET_AREA', target_area).replace('QUERY_TABLE', data_table)

        i = 1  # �к�
        for row in cursor.execute(sql):
            i += 1
            hour = i-2
            sheet.Cells(i, 3).Value = row[0]  # ��������
            sheet.Cells(i, 6).Value = positive_intensity_hours = row[1] if row[1] is not None else 0  # ����ǿ��
            sum_hour_dict[hour] += row[0]
            Stats_of_Hour[hour].append(row[0])
            Stats_of_Hour[hour].append(positive_intensity_hours)
            Stats_of_Hour[hour].append(sum_hour_dict[hour])

        query_results['Stats_of_Hour'] = Stats_of_Hour

        # **********����ǿ�ȷֲ�**************
        sql = """
        SELECT count(*) AS ��������,0 AS ��߽�,5 AS �ұ߽�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=0 AND Abs(Intensity)<5
        union SELECT count(*) AS ��������,5,10
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=5 AND Abs(Intensity)<10
        union SELECT count(*) AS ��������,10,15
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=10 AND Abs(Intensity)<15
        union SELECT count(*) AS ��������,15,20
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=15 AND Abs(Intensity)<20
        union SELECT count(*) AS ��������,20,25
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=20 AND Abs(Intensity)<25
        union SELECT count(*) AS ��������,25,30
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=25 AND Abs(Intensity)<30
        union SELECT count(*) AS ��������,30,35
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=30 AND Abs(Intensity)<35
        union SELECT count(*) AS ��������,35,40
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=35 AND Abs(Intensity)<40
        union SELECT count(*) AS ��������,40,45
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=40 AND Abs(Intensity)<45
        union SELECT count(*) AS ��������,45,50
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=45 AND Abs(Intensity)<50
        union SELECT count(*) AS ��������,50,55
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=50 AND Abs(Intensity)<55
        union SELECT count(*) AS ��������,55,60
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=55 AND Abs(Intensity)<60
        union SELECT count(*) AS ��������,60,65
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=60 AND Abs(Intensity)<65
        union SELECT count(*) AS ��������,65,70
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=65 AND Abs(Intensity)<70
        union SELECT count(*) AS ��������,70,75
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=70 AND Abs(Intensity)<75
        union SELECT count(*) AS ��������,75,80
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=75 AND Abs(Intensity)<80
        union SELECT count(*) AS ��������,80,85
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=80 AND Abs(Intensity)<85
        union SELECT count(*) AS ��������,85,90
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=85 AND Abs(Intensity)<90
        union SELECT count(*) AS ��������,90,95
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=90 AND Abs(Intensity)<95
        union SELECT count(*) AS ��������,95,100
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=95 AND Abs(Intensity)<100
        union SELECT count(*) AS ��������,100,150
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=100 AND Abs(Intensity)<150
        union SELECT count(*) AS ��������,150,200
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=150 AND Abs(Intensity)<200
        union SELECT count(*) AS ��������,200,250
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=200 AND Abs(Intensity)<250
        union SELECT count(*) AS ��������,250,300
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=250 AND Abs(Intensity)<300
        UNION SELECT count(*) AS ��������,300,1000
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity<0 AND Abs(Intensity)>=300
        ORDER BY ��߽�
        """.replace('TARGET_AREA', target_area).replace("QUERY_TABLE", data_table)

        sheet = workbook.Worksheets(u'ǿ�ȷֲ�ͳ��')
        i = 1  # �к�
        for row in cursor.execute(sql):
            i += 1
            sheet.Cells(i, 3).Value = row[0]  # ��������

        # ***********����ǿ�ȷֲ�************
        sql = """
        SELECT count(*) AS ��������,0 AS ��߽�,5 AS �ұ߽�
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=0 AND Intensity<5
        union SELECT count(*) AS ��������,5,10
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=5 AND Intensity<10
        union SELECT count(*) AS ��������,10,15
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=10 AND Intensity<15
        union SELECT count(*) AS ��������,15,20
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=15 AND Intensity<20
        union SELECT count(*) AS ��������,20,25
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=20 AND Intensity<25
        union SELECT count(*) AS ��������,25,30
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=25 AND Intensity<30
        union SELECT count(*) AS ��������,30,35
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=30 AND Intensity<35
        union SELECT count(*) AS ��������,35,40
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=35 AND Intensity<40
        union SELECT count(*) AS ��������,40,45
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=40 AND Intensity<45
        union SELECT count(*) AS ��������,45,50
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=45 AND Intensity<50
        union SELECT count(*) AS ��������,50,55
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=50 AND Intensity<55
        union SELECT count(*) AS ��������,55,60
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=55 AND Intensity<60
        union SELECT count(*) AS ��������,60,65
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=60 AND Intensity<65
        union SELECT count(*) AS ��������,65,70
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=65 AND Intensity<70
        union SELECT count(*) AS ��������,70,75
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=70 AND Intensity<75
        union SELECT count(*) AS ��������,75,80
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=75 AND Intensity<80
        union SELECT count(*) AS ��������,80,85
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=80 AND Intensity<85
        union SELECT count(*) AS ��������,85,90
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=85 AND Intensity<90
        union SELECT count(*) AS ��������,90,95
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=90 AND Intensity<95
        union SELECT count(*) AS ��������,95,100
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=95 AND Intensity<100
        union SELECT count(*) AS ��������,100,150
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=100 AND Intensity<150
        union SELECT count(*) AS ��������,150,200
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=150 AND Intensity<200
        union SELECT count(*) AS ��������,200,250
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=200 AND Intensity<250
        union SELECT count(*) AS ��������,250,300
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=250 AND Intensity<300
        UNION SELECT count(*) AS ��������,300,1000
        FROM QUERY_TABLE
        WHERE Region='TARGET_AREA' AND Intensity>=300
        ORDER BY ��߽�
        """.replace('TARGET_AREA', target_area).replace("QUERY_TABLE", data_table)

        i = 1  # �к�
        for row in cursor.execute(sql):
            i += 1
            sheet.Cells(i, 4).Value = row[0]  # ��������

        f = open(os.path.join(workspath, 'query_results.pkl'), 'wb')
        pickle.dump(query_results, f, 2)#���ܲ�����ߵ�Э�飬������python2.7.8���޷����ؽ���
        f.close()


    finally:
        db.close()  # �ر���������
        workbook.Save()  # ����EXCEL������
        workbook.Close()  # �رչ������ļ�
        excel.Quit()  # �ر�EXCELӦ�ó���


if __name__ == "__main__":

    (year, province, target_area,cwd) = sys.argv[1:]
    sqlQuery(year, province, target_area,cwd)

    # year = u"2015��"
    # province = u'�㽭'
    # target_area = u"������"
    #
    # cwd = os.getcwd()
    # start = time.clock()
    # # ***********************���Գ���*********************************"
    # sqlQuery(year, province, target_area, cwd)
    # # ***********************���Գ���*********************************"
    # end = time.clock()
    # elapsed = end - start
    # print("Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000))
