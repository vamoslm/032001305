import webbrowser
import pyecharts.options as opts
import xlrd
from pyecharts.charts import Line, Map




def show_new_confirmed_data(NEW_CONFIRMED_CASES_DATA_EXCEL, NEW_SPECIAL_ZONES_DATA_EXCEL):
    data = xlrd.open_workbook(NEW_CONFIRMED_CASES_DATA_EXCEL)  # 打开本地excel表格
    table = data.sheet_by_index(0)  # 拿出表格的第一个sheet

    province = ['湖北', '北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江', '上海', '江苏', '浙江', '安徽', '福建', '江西', '山东',
                '河南', '湖南', '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海', '宁夏', '新疆']
    infections = []
    i = 2
    # 循环输出每行内容
    for item in province:
        infections.append(int(table.cell(1, i).value))
        i += 1

    # 处理港澳台数据
    special_data = xlrd.open_workbook(NEW_SPECIAL_ZONES_DATA_EXCEL)
    sheet1 = special_data.sheet_by_index(0)
    province.append('台湾')
    infections.append(int(sheet1.cell(1, 1).value - sheet1.cell(2, 1).value))
    province.append('香港')
    infections.append(int(sheet1.cell(1, 2).value - sheet1.cell(2, 2).value))
    province.append('澳门')
    infections.append(int(sheet1.cell(1, 3).value - sheet1.cell(2, 3).value))

    list1 = [[province[i], infections[i]] for i in range(len(province))]
    map1 = Map(init_opts=opts.InitOpts(height="800px", width="1250px"))
    map1.set_global_opts(visualmap_opts=opts.VisualMapOpts(is_piecewise=True,
        pieces=[ {"max": 0,  "min": 0, "label": "0", "color":"#ffffff"},
                 {"max": 10, "min": 1, "label": "1-10", "color": "#00FFFF"},
                 {"max": 20, "min": 11, "label": "11-20", "color": "#0dd6e1"},
                 {"max": 30, "min": 21, "label": "21-30", "color": "#0d6ee1"},
                 {"max": 40, "min": 31, "label": "31-40", "color": "#e1d10d"},
                 {"max": 50, "min": 41, "label": "41-50", "color": "#e1600d"},
                 {"max": 100, "min": 51, "label": "51-100", "color": "#c75113"},
                 { "min": 101, "label": ">100", "color": "#c10303"}
                ]),toolbox_opts=opts.ToolboxOpts(is_show=True),legend_opts=opts.LegendOpts(is_show=False),
                title_opts=opts.TitleOpts(title=f'中国{str(table.cell(1,0).value)}各省市新增确诊人数',pos_left='40%', pos_top='10'))
    # 标签名称显示，默认为True
    map1.set_series_opts(label_opts=opts.LabelOpts(is_show=True, color="black"))
    map1.add(f'{str(table.cell(1,0).value)}',list1)
    return map1
    # map.render(f'中国近一日来各省市新增确诊人数.html')
    # # 在浏览器中自动打开
    # webbrowser.open_new_tab(f'中国近一日来各省市新增确诊人数.html')


def show_year_confirmed_data(NEW_CONFIRMED_CASES_DATA_EXCEL):
    data = xlrd.open_workbook(NEW_CONFIRMED_CASES_DATA_EXCEL)  # 打开本地excel表格
    table = data.sheet_by_index(0)  # 拿出表格的第一个sheet
    infections = []
    date = []
    col_count = 1
    while col_count <= 365:
        date.append(str(table.cell(365 - col_count, 0).value))
        infections.append(table.cell(365 - col_count, 1).value)
        col_count += 1
    line = Line(init_opts=opts.InitOpts(width="10000px", height="600px"))  # 创建一个柱状图对象
    line.add_xaxis(date)  # 设置x轴
    line.add_yaxis(f'中国近一年来每日新增确诊人数', infections ,is_smooth=True,is_hover_animation=True)  # 设置y轴的参数
    line.render(f'中国近一年来每日新增确诊人数.html')  # 输出html文件来显示柱状图
    # 在浏览器中自动打开
    webbrowser.open_new_tab(f'中国近一年来每日新增确诊人数.html')


def show_twenty_days_data(NEW_CONFIRMED_CASES_DATA_EXCEL):
    data = xlrd.open_workbook(NEW_CONFIRMED_CASES_DATA_EXCEL)  # 打开本地excel表格
    table = data.sheet_by_index(0)  # 拿出表格的第一个sheet
    twenty_days = 21
    province = ['湖北', '北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江', '上海', '江苏', '浙江',
                '安徽', '福建', '江西', '山东',
                '河南', '湖南', '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海',
                '宁夏', '新疆']
    line = Line(init_opts=opts.InitOpts(width="1500px", height="800px"))  # 创建一个柱状图对象
    date = []
    column = 2
    while twenty_days >= 1:
        date.append(str(table.cell(twenty_days,0).value))
        twenty_days -= 1
    line.add_xaxis(date)  # 设置x轴
    for item in province:
        twenty_days = 21
        infection = []
        while twenty_days >= 1:
            infection.append(int(table.cell(twenty_days,column).value))
            twenty_days -= 1
        # 省份多，折线杂，所以将每个省份近二十天来的新增确诊人数最大值标注出来
        line.add_yaxis(f'{str(item)}',infection,is_smooth=True,is_hover_animation=True,
            markpoint_opts=opts.MarkPointOpts(data=[opts.MarkPointItem(type_="max", name=f"{item}近二十日新增人数最大值")]))
        column += 1
    line.set_global_opts(legend_opts=opts.LegendOpts(type_="scroll", pos_left="left", orient="vertical"), toolbox_opts=opts.ToolboxOpts(is_show=True),
                         title_opts=opts.TitleOpts(title=f'中国近二十天来各省市新增确诊人数',pos_left='40%', pos_top='10'))
    return line
    # line.render('中国近二十天来各省市新增确诊人数.html')
    # webbrowser.open_new_tab('中国近二十天来各省市新增确诊人数.html')


def show_twenty_days_asymptomaitc(NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL):
    data = xlrd.open_workbook(NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL)  # 打开本地excel表格
    table = data.sheet_by_index(0)  # 拿出表格的第一个sheet
    twenty_days = 21
    province = ['湖北', '北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江', '上海', '江苏', '浙江',
                '安徽', '福建', '江西', '山东',
                '河南', '湖南', '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海',
                '宁夏', '新疆']
    line = Line(init_opts=opts.InitOpts(width="1500px", height="800px"))  # 创建一个柱状图对象
    date = []
    column = 2
    while twenty_days >= 1:
        date.append(str(table.cell(twenty_days, 0).value))
        twenty_days -= 1
    line.add_xaxis(date)  # 设置x轴
    for item in province:
        twenty_days = 21
        infection = []
        while twenty_days >= 1:
            infection.append(int(table.cell(twenty_days, column).value))
            twenty_days -= 1
        # 省份多，折线杂，所以将每个省份近二十天来的新增无症状感染者人数最大值标注出来
        line.add_yaxis(f'{str(item)}', infection, is_smooth=True, is_hover_animation=True,
                       markpoint_opts=opts.MarkPointOpts(
                           data=[opts.MarkPointItem(type_="max", name=f"{item}近二十日新增无症状感染者人数最大值")]))
        column += 1
    line.set_global_opts(legend_opts=opts.LegendOpts(type_="scroll", pos_left="left", orient="vertical"),
                         toolbox_opts=opts.ToolboxOpts(is_show=True),
                         title_opts=opts.TitleOpts(title=f'中国近二十天来各省市新增无症状感染者人数', pos_left='40%',
                                                   pos_top='10'))
    return line
    # line.render('中国近二十天来各省市新增无症状感染者人数.html')
    # webbrowser.open_new_tab('中国近二十天来各省市新增无症状感染者人数.html')

def hotpoint(NEW_CONFIRMED_CASES_DATA_EXCEL):
    data = xlrd.open_workbook(NEW_CONFIRMED_CASES_DATA_EXCEL)  # 打开本地excel表格
    table = data.sheet_by_index(0)  # 拿出表格的第一个sheet
    print(f'{str(table.cell(1,0).value)}热点事件：')
    province = ['湖北', '北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江', '上海', '江苏', '浙江',
                '安徽', '福建', '江西', '山东',
                '河南', '湖南', '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海',
                '宁夏', '新疆']
    # 判断哪些省市近七天来第一天出现新增确诊病例
    sevendays_cites = []
    infections = []
    city_counts = 0
    column = 2
    for item in province:
        mark = 0
        row = 2
        counts = 0
        while row <= 8:
            if int(table.cell(1,column).value) == 0:
                mark = 1
                break
            elif int(table.cell(row,column).value) != 0:
                counts += int(table.cell(row, column).value)
                mark = 1
                break
            row += 1
        if mark == 0:
            sevendays_cites.append(f'{item}')
            infections.append(column)
            city_counts += 1
        column += 1
    if city_counts == 0:
        print('今日未有七天内首次出现新增确诊病例的城市')
    else:
        print(f'一共有{city_counts}座城市七天内首次出现新增确诊病例：')
        i = 0
        while i < len(sevendays_cites):
            print(f'{sevendays_cites[i]}：{int(table.cell(1,infections[i]))}例    ',end='')
            i += 1
    # 判断哪些省市连续七天都有新增确诊病例
    city_counts = 0
    column = 2
    sevendays_cites = []
    infections = []
    for item in province:
        mark = 0
        row = 1
        counts = 0
        while row <= 7:
            counts += int(table.cell(row,column).value)
            if int(table.cell(row,column).value) == 0:
                mark = 1
                break
            row += 1
        if mark == 0:
            city_counts += 1
            sevendays_cites.append(f'{item}')
            infections.append(counts)
        column += 1
    if city_counts == 0:
        print('今日未有连续七天出现新增确诊病例的城市')
    else:
        print(f'一共有{city_counts}座城市连续七天都有新增确诊病例：')
        i = 0
        while i < len(sevendays_cites):
            print(f'{sevendays_cites[i]}七天累计{infections[i]}例')
            i += 1