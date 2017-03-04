# -*- coding: utf-8 -*-

from win32com.client import DispatchEx, constants
from win32com.client.gencache import EnsureDispatch


def docProcess():


    EnsureDispatch('Word.Application')
    word = DispatchEx('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(''.join([cwd, u'/Data/公报模板.docx']))
    try:

        #**********分段落处理***********
        #todo sum_region_lastyear
        sum_region_lastyear = 22122
        day_region = 40
        if 0.05<=(sum_region-sum_region_lastyear)/float(sum_region)<0.1:
            compare_with_lastyear = u'略有增多'
        elif -0.1<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.05:
            compare_with_lastyear = u'略有较少'
        elif 0.1<=(sum_region-sum_region_lastyear)/float(sum_region)<0.5:
            compare_with_lastyear = u'有所增多'
        elif -0.5<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.1:
            compare_with_lastyear = u'有所较少'
        elif 0.5<=(sum_region-sum_region_lastyear)/float(sum_region)<0.9:
            compare_with_lastyear = u'增幅较大'
        elif -0.9<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.5:
            compare_with_lastyear = u'减幅较大'
        elif 0.9<=(sum_region-sum_region_lastyear)/float(sum_region):
            compare_with_lastyear = u'大幅增多'
        elif (sum_region-sum_region_lastyear)/float(sum_region)<=-0.9:
            compare_with_lastyear = u'大幅减少'
        else:
            compare_with_lastyear = u'基本持平'

        if density_region>density_province:
            compare_with_province = u'高于'
        else:
            compare_with_province = u'低于'

        p24 = u'%d年我市共发生地闪%d次，平均地闪密度%.2f次/km²，平均雷暴日%d天（见表1-1）。\
与上年的地闪%d次相比，%s。从时间分布来看，地闪主要集中在%d、%d、%d月，\
三个月地闪占全年总地闪次数的%.2f%%。从空间分布来看，%s发生地闪次数最多，%s最少。\
全市地闪平均密度%s全省平均的%.2f次/km²，在全省各市中%s闪次数排第%d位，\
地闪平均密度排第%d位（见表1-2）。\n'% (year,sum_region,density_region,day_region,
                                sum_region_lastyear, compare_with_lastyear, max_months[0],max_months[1],max_months[2],
                                max_months_percent,sum_max_county,sum_min_county,compare_with_province,density_province,
                                target_area,sum_rank_in_province,density_rank_in_province)


        rng  = doc.Paragraphs(24).Range
        paragraphFormat = rng.ParagraphFormat
        paragraphFont = rng.Font
        rng.Text = p24


        p25 = u'据不完全统计，2016年全市因雷电引发的灾害共148起，无人员伤亡事故。\
造成直接经济损失达7788.04万元，间接经济损失677.42万元。\n'

        rng  = doc.Paragraphs(25).Range
        rng.Text = p25
        rng.ParagraphFormat = paragraphFormat

        p109= u'从地区统计来看，地区分布相对不均，%s地闪次数最多，共%d次，%s最少，只有%d次，\
两者分别占全市总地闪数的%.2f%%和%.2f%%。从平均密度统计来看，%s密度最高，为%.2f次/km²，\
%s最低，为%.2f次/km²（见表1-1）。\n'%(sum_max_county,sum_max_region,sum_min_county,sum_min_region,
                                            max_region_percent,min_region_percent,
                                            density_max_county,density_max_region,
                                            density_min_county,density_min_region)

        rng  = doc.Paragraphs(109).Range
        rng.Text = p109
        rng.ParagraphFormat = paragraphFormat

        p110 = u'从地闪密度空间分布图上（见图1-1）可以看出，诸暨西北部、嵊州和诸暨交界区域地闪密度较高，\
最高超过5次/km²。新昌东部有部分地区，地闪密度超过3次/km²，全市大部分地区地闪密度小于2次/km²。\n'

        rng  = doc.Paragraphs(110).Range
        rng.Text = p110
        rng.ParagraphFormat = paragraphFormat

        p113= u'现行国家标准所引用的雷暴日指人工观测（测站周围约15km半径域面）有雷暴天数的多年平均。\
根据我省闪电定位监测资料推算（以15km为间隔，分别统计各点15km半径范围内的雷暴日，再插值推算），\
2016年全市地闪雷暴日平均43天，最低为29天，最高67天。空间分布上来看，北部平原地区雷暴日较少，\
西南大部和东南部分区域雷暴天数较多（见图1-2）。\n'

        rng  = doc.Paragraphs(113).Range
        rng.Text = p113
        rng.ParagraphFormat = paragraphFormat

        if len(months_zero) ==0:
            months_zero_description = u''
        elif len(months_zero) ==1:
            months_zero_description = u'%d月未监测到地闪，'%months_zero[0]
        elif len(months_zero) ==2:
            months_zero_description = u'%d月和%d月未监测到地闪，'%(months_zero[0],months_zero[1])
        else:
            s =  u'月、'.join(map(lambda d:str(d),months_zero[:len(months_zero)-1]))
            months_zero_description = u''.join([s,u'月和%d月都未监测到地闪，'%months_zero[-1]])
        p115 =u'%d年%s雷电初日为%d月%d日。从分月统计来看，地闪次数随月份呈现近似正态分布特征，%s\
地闪次数峰值出现在%d月，%d、%d、%d月是雷暴高发的月份，三个月地闪次数占总数的%.2f%%。\
正、负地闪平均强度的峰值分别在%d月和%d月，其他月份波动平缓(见图1-3) 。\n'%(year,target_area,
                        first_date.month,first_date.day,months_zero_description,
                        max_month_region,max_months[0],max_months[1],max_months[2],max_months_percent,
                        peak_month_positive_intensity,peak_month_negative_intensity)

        rng  = doc.Paragraphs(115).Range
        rng.Text = p115
        rng.ParagraphFormat = paragraphFormat

        p118 = u'从分时段统计来看，地闪次数峰值出现在第18个时段（17:00-18:00），地闪主要集中在午后两点到晚上九点，\
七个时段内的地闪次数占总数的d%。地闪平均强度随时间呈波状起伏特征，但总体波动不大。\
正闪平均强度峰值在第7个时段（7:00-8:00），负闪平均强度峰值在第11个时段（11:00-12:00）(见图1-4)。\n'

        rng  = doc.Paragraphs(118).Range
        rng.Text = p118
        rng.ParagraphFormat = paragraphFormat
        rng.Font = paragraphFont

        p122 = u'由正、负地闪强度分布图可见，地闪次数随地闪强度呈近似正态分布特征。正地闪主要集中在5-60kA内（见图1-5），\
该区间内正地闪次数约占总地闪的87.20%，负地闪主要分布在5-60kA内（见图1-6），\
该区间内负地闪次数约占总负地闪的91.46%。\n'
        rng  = doc.Paragraphs(122).Range
        rng.Text = p122
        rng.ParagraphFormat = paragraphFormat


    finally:
        doc.Save()
        doc.Close()
        word.Quit()