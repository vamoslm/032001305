import webbrowser
import xlwt
import showdata
import spider
import xlrd
from pyecharts.charts import Page


TOTAL_HOME_PAGES = 41
HOMEPAGE_EXCEL = 'C:/Users/vamos/Desktop/homepage_url.xlsx'
SUBPAGE_EXCEL = 'C:/Users/vamos/Desktop/subpage_url.xlsx'
NEW_CONFIRMED_CASES_DATA_EXCEL = 'C:/Users/vamos/Desktop/大陆地区新增确诊人数.xlsx'
NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL = 'C:/Users/vamos/Desktop/大陆地区新增无症状感染者人数.xlsx'
NEW_SPECIAL_ZONES_DATA_EXCEL = 'C:/Users/vamos/Desktop/港澳台累计确诊人数.xlsx'


# 爬取一级页面网址并创建和导入相应的excel
spider.sethomepage(HOMEPAGE_EXCEL)

# 遍历一级页面网址，获取对应源码筛选出二级界面并创建相应的excel
page_count = 1
row_count = 1  # 记录excel的行数
# 建立excel存储二级网站
wb = xlwt.Workbook()
ws = wb.add_sheet("Subpage Sheet")  # 创建sheet
ws.write(0, 0, "二级页面网址：")  # 设置表单标题
while page_count <= TOTAL_HOME_PAGES:
    home_url = spider.gethomepage(page_count, HOMEPAGE_EXCEL)   # 根据行数提取网址
    home_page = spider.getdatapage(home_url, page_count)     # 爬取一级源码
    spider.setsubpage(ws, wb, home_page, page_count, SUBPAGE_EXCEL)  # 利用a标签提取大致网址并利用href标签提取具体网站导入新创建的excel当中
    page_count += 1

# 返回行数得知一共有几条二级页面网址
subpage_counts = spider.getsubpagecount(SUBPAGE_EXCEL)
print(f'=====共有{subpage_counts}条子网址=====')

# # 遍历网址，取出数据
spider.creatdataexcel(NEW_CONFIRMED_CASES_DATA_EXCEL, NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL,
                      NEW_SPECIAL_ZONES_DATA_EXCEL)  # 创建用来统计数据的excel
page_count = 1
row_count = page_count
wb = xlrd.open_workbook(SUBPAGE_EXCEL)  # 读取excel
# 根据索引获取sheet内容
sheet1_content = wb.sheet_by_index(0)
while page_count <= subpage_counts:
    data = spider.getdata(sheet1_content, page_count)  # 获取文本数据
    if data is None:
        page_count += 1
        continue
    else:
        spider.new_confirmed_cases(data, row_count, NEW_CONFIRMED_CASES_DATA_EXCEL)  # 获取新增确诊人数数据
        spider.new_asymptomatic_infections(data, row_count, NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL)
        spider.specialzones(data, row_count, NEW_SPECIAL_ZONES_DATA_EXCEL)
        page_count += 1

# 生成中国近一年来每日新增确诊人数并展示
showdata.show_year_confirmed_data(NEW_CONFIRMED_CASES_DATA_EXCEL)
# 生成中国近一个月来各省市新增确诊人数
line1 = showdata.show_twenty_days_data(NEW_CONFIRMED_CASES_DATA_EXCEL)
# 生成中国近一个月来各省市新增无症状感染者人数
line2 = showdata.show_twenty_days_asymptomaitc(NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL)
# 生成中国近一日各省市新增确诊人数
map1 = showdata.show_new_confirmed_data(NEW_CONFIRMED_CASES_DATA_EXCEL, NEW_SPECIAL_ZONES_DATA_EXCEL)
# 将多张图合成一个html进行展示
page = Page()
page.add(map1, line1, line2)
page.render('数据展示.html')
webbrowser.open_new_tab('数据展示.html')

# 展示热点事件
# 1.哪些省市近七天来第一天出现新增确诊病例
# 2.哪些省市连续七天都有新增确诊病例
showdata.hotpoint(NEW_CONFIRMED_CASES_DATA_EXCEL)
