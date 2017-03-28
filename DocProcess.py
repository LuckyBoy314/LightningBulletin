# -*- coding: utf-8 -*-

from win32com.client import DispatchEx, constants
from win32com.client.gencache import EnsureDispatch
import shutil, os, time
import pickle


def docProcess(year, province, target_area, cwd):
    # 打开word文档
    EnsureDispatch('Word.Application')
    word = DispatchEx('Word.Application')
    word.Visible = False

    doc_origin = ''.join([cwd, u'/data/', u'公报文档模板.docx'])
    doc = ''.join([cwd, u'/temp/', province, '/', year, '/', target_area, '.gdb/',
                   year, target_area, u'公报文档.docx'])

    #if not os.path.exists(doc):
    shutil.copy2(doc_origin, doc)

    doc = word.Documents.Open(doc)

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

    # ******替换图片*********
    word.Selection.Find.Execute(FindText=u'图1-1 地闪密度空间分布图', Wrap=1)
    word.Selection.MoveLeft(Count = 3)
    doc.InlineShapes(2).Delete()
    densityPic = ''.join([cwd, u'/temp/', province, u'/',year, u'/', target_area, u'.gdb/',
                        year, target_area, u'闪电密度空间分布.png'])
    word.Selection.InlineShapes.AddPicture(densityPic)

    word.Selection.Find.Execute(FindText=u'图1-2 地闪雷暴日空间分布图', Wrap=1)
    word.Selection.MoveLeft(Count = 3)
    doc.InlineShapes(3).Delete()
    dayPic = ''.join([cwd, u'/temp/', province, u'/',year, u'/', target_area, u'.gdb/',
                        year, target_area, u'地闪雷暴日空间分布.png'])
    word.Selection.InlineShapes.AddPicture(dayPic)

    # 读取统计参数
    query_results_path = ''.join([cwd, u"/temp/", province, '/', year, '/', target_area, '.gdb/', 'query_results.pkl'])
    # 读取查询参数
    f = open(query_results_path, 'rb')
    query_results = pickle.load(f)
    f.close()

    sum_target_area = query_results['Sum_target_area']  # 本地区地闪总数
    sum_rank_in_province = query_results['Sum_rank_in_province']  # 本地区地闪总数在全省排名
    density_province = query_results['Density_province']  # 全省密度
    density_target_area = query_results['Density_target_area']  # 本地区密度
    density_rank_in_province = query_results['Density_rank_in_province']  # 本地区密度在全省排名
    stats_of_Region = query_results['Stats_of_Region']  # 分县市统计
    day_target_area = stats_of_Region[u'总计'][4]

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

    try:
        # 处理与去年的对比情况
        last_year_query_results_path = query_results_path.replace(year, str(int(year[:-1]) - 1) + u'年')
        if os.path.exists(last_year_query_results_path):
            f = open(last_year_query_results_path, 'rb')
            last_year_query_results = pickle.load(f)
            f.close()

            sum_target_area_last_year = last_year_query_results['Sum_target_area']

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

        p24 = u'%s我市共发生地闪%d次，平均地闪密度%.2f次/km²，平均雷暴日%d天（见表1-1）。%s\
从时间分布来看，地闪主要集中在%d、%d、%d月，三个月地闪占全年总地闪次数的%.2f%%。从空间分布来看，%s发生地闪次数最多，%s最少。\
全市地闪平均密度%s全省平均的%.2f次/km²，在全省各市中%s闪次数排第%d位，地闪平均密度排第%d位（见表1-2）。\n' % (
            year, sum_target_area, density_target_area, day_target_area, compare_with_last_year,
            max_three_months[0], max_three_months[1], max_three_months[2], max_months_percent,
            sum_max_county_name, sum_min_county_name, compare_with_province, density_province,
            target_area, sum_rank_in_province, density_rank_in_province)

        rng = doc.Paragraphs(24).Range
        paragraphFormat = rng.ParagraphFormat
        paragraphFont = rng.Font
        rng.Text = p24

        p25 = u'据不完全统计，2016年全市因雷电引发的灾害共148起，无人员伤亡事故。\
造成直接经济损失达7788.04万元，间接经济损失677.42万元。\n'

        rng = doc.Paragraphs(25).Range
        rng.Text = p25
        rng.ParagraphFormat = paragraphFormat

        p109 = u'从地区统计来看，地区分布相对不均，%s地闪次数最多，共%d次，%s最少，只有%d次，\
两者分别占全市总地闪数的%.2f%%和%.2f%%。从平均密度统计来看，%s密度最高，为%.2f次/km²，\
%s最低，为%.2f次/km²（见表1-1）。\n' % (sum_max_county_name, sum_max_in_region, sum_min_county_name, sum_min_in_region,
                              max_region_percent, min_region_percent,
                              density_max_county_name, density_max_in_region,
                              density_min_county_name, density_min_in_region)

        rng = doc.Paragraphs(109).Range
        rng.Text = p109
        rng.ParagraphFormat = paragraphFormat

        p110 = u'从地闪密度空间分布图上（见图1-1）可以看出，诸暨西北部、嵊州和诸暨交界区域地闪密度较高，\
最高超过5次/km²。新昌东部有部分地区，地闪密度超过3次/km²，全市大部分地区地闪密度小于2次/km²。\n'

        rng = doc.Paragraphs(110).Range
        rng.Text = p110
        rng.ParagraphFormat = paragraphFormat

        p113 = u'现行国家标准所引用的雷暴日指人工观测（测站周围约15km半径域面）有雷暴天数的多年平均。\
根据我省闪电定位监测资料推算（以15km为间隔，分别统计各点15km半径范围内的雷暴日，再插值推算），\
2016年全市地闪雷暴日平均43天，最低为29天，最高67天。空间分布上来看，北部平原地区雷暴日较少，\
西南大部和东南部分区域雷暴天数较多（见图1-2）。\n'

        rng = doc.Paragraphs(113).Range
        rng.Text = p113
        rng.ParagraphFormat = paragraphFormat

        if len(months_zero) == 0:
            months_zero_description = u''
        elif len(months_zero) == 1:
            months_zero_description = u'%d月未监测到地闪，' % months_zero[0]
        elif len(months_zero) == 2:
            months_zero_description = u'%d月和%d月未监测到地闪，' % (months_zero[0], months_zero[1])
        else:
            s = u'月、'.join(map(lambda d: str(d), months_zero[:len(months_zero) - 1]))
            months_zero_description = u''.join([s, u'月和%d月都未监测到地闪，' % months_zero[-1]])

        p115 = u'%s%s雷电初日为%s。从分月统计来看，地闪次数随月份呈现近似正态分布特征，%s\
地闪次数峰值出现在%d月，%d、%d、%d月是雷暴高发的月份，三个月地闪次数占总数的%.2f%%。\
正、负地闪平均强度的峰值分别在%d月和%d月，其他月份波动平缓(见图1-3) 。\n' % (year, target_area,
                                               first_date, months_zero_description,
                                               max_month_region, max_three_months[0], max_three_months[1],
                                               max_three_months[2],
                                               max_months_percent,
                                               peak_month_positive_intensity, peak_month_negative_intensity)

        rng = doc.Paragraphs(115).Range
        rng.Text = p115
        rng.ParagraphFormat = paragraphFormat

        p118 = u'从分时段统计来看，地闪次数峰值出现在第18个时段（17:00-18:00），地闪主要集中在午后两点到晚上九点，\
七个时段内的地闪次数占总数的d%。地闪平均强度随时间呈波状起伏特征，但总体波动不大。\
正闪平均强度峰值在第7个时段（7:00-8:00），负闪平均强度峰值在第11个时段（11:00-12:00）(见图1-4)。\n'

        rng = doc.Paragraphs(118).Range
        rng.Text = p118
        rng.ParagraphFormat = paragraphFormat
        rng.Font = paragraphFont

        p122 = u'由正、负地闪强度分布图可见，地闪次数随地闪强度呈近似正态分布特征。正地闪主要集中在5-60kA内（见图1-5），\
该区间内正地闪次数约占总地闪的87.20%，负地闪主要分布在5-60kA内（见图1-6），\
该区间内负地闪次数约占总负地闪的91.46%。\n'
        rng = doc.Paragraphs(122).Range
        rng.Text = p122
        rng.ParagraphFormat = paragraphFormat


    finally:
        doc.Save()
        doc.Close()
        word.Quit()


if __name__ == "__main__":
    # (year, province, target_area,cwd) = sys.argv[1:]
    # sqlQuery(year, province, target_area,cwd)

    year = u"2016年"
    province = u'河南'
    target_area = u"新乡市"

    cwd = os.getcwd()
    start = time.clock()
    # ***********************测试程序*********************************"
    docProcess(year, province, target_area, cwd)
    # ***********************测试程序*********************************"
    end = time.clock()
    elapsed = end - start
    print("Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000))
