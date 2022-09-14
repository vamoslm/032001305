import time
import requests
import xlwt
import xlrd
import pyecharts.options as opts
from pyecharts.charts import Line
from xlutils.copy import copy
import re
import webbrowser
from bs4 import BeautifulSoup


headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.27'}

def sethomepage(HOMEPAGE_EXCEL):
    print('=====开始爬取一级页面网址，并导入excel',end='')
    # 导出url到excel当中
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Homepage Sheet")  # 创建sheet
    ws.write(0, 0, "一级页面网址：")
    row_count = 1  # 记录excel的行数
    page_count = 1;

    while page_count <= 41:
        if(page_count == 1):
            ws.write(row_count,0,'http://www.nhc.gov.cn/xcs/yqtb/list_gzbd.shtml')
        else:
            ws.write(row_count,0,f'http://www.nhc.gov.cn/xcs/yqtb/list_gzbd_{str(page_count)}.shtml')
        page_count += 1
        row_count += 1

    wb.save(HOMEPAGE_EXCEL) #保存excel
    print('成功=====')


def gethomepage(page_count,HOMEPAGE_EXCEL):
    wb = xlrd.open_workbook(HOMEPAGE_EXCEL) # 读取excel
    # 根据索引获取sheet内容
    sheet1_content = wb.sheet_by_index(0);
    # 获取行当中的数据
    rows = sheet1_content.row_values(page_count)
    # print(rows)
    # print(type(rows))  # row类型为list
    # 将其转换为str类型，以便使用
    url_str = rows[0]
    return url_str
    # print(url_str)
    # print(type(url_str))


def getdatapage(url,page_count):
    print(url)
    while True:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            print(f"page{page_count} sucessed")
            break
        else:
            print(f"page{page_count} failed  ",end='')
            time.sleep(1)
    data_page = response.content.decode()
    return data_page
    # print(data_page)

row_count = 1  # 记录excel的行数
wb = xlwt.Workbook()
ws = wb.add_sheet("Subpage Sheet")  # 创建sheet
ws.write(0, 0, "二级页面网址：")  # 设置标题
def setsubpage(home_page, page_count,SUBPAGE_EXCEL):
    print(f'=====爬取第{page_count}页二级页面网址成功=====')
    soup = BeautifulSoup(home_page, 'lxml')
    arr = soup.find_all('a')  # 利用a标签提取各日疫情通报网站
    # print(arr)
    for term in arr:
        url = term['href'] # 提取href内容
        sub_url = str('http://www.nhc.gov.cn/'+url)
        # print(sub_url)
        # print(type(sub_url))
        # 筛选是疫情通报的网址
        if sub_url.__contains__('yqtb'):
            global row_count # 声明全局变量
            ws.write(row_count,0,row_count)
            ws.write(row_count,1,sub_url)
            row_count += 1
    wb.save(SUBPAGE_EXCEL)


def getsubpagecount(SUBPAGE_EXCEL):
    wb = xlrd.open_workbook(SUBPAGE_EXCEL) # 读取excel
    # 根据索引获取sheet内容
    sheet1_content = wb.sheet_by_index(0)
    return sheet1_content.nrows - 1


def getdata(sheet1_content,page_count):
    print(f'=====读取第{page_count}条子页面数据=====')
    # 获取行当中的数据
    url = sheet1_content.row_values(page_count,1)
    # 将其转换为str类型，以便使用
    url_str = url[0]
    # print(url_str)
    # print(type(url_str))
    data_page = getdatapage(url_str,page_count)
    # print(datapage)
    # 使用BeautifulSoup提取疫情数据
    soup = BeautifulSoup(data_page, 'lxml')
    p_all = soup.find_all('p')  # 利用p标签提取疫情数据
    # print(p_all)
    tit = soup.find_all(class_='tit')
    print(tit)
    if tit[0].text.__contains__('国务院联防联控机制'):
        return None
    elif tit[0].text.__contains__('武汉市新冠'):
        return None
    else:
        return p_all
    # 获取标签中文本内容
    # print(p_all[0].text)
    # print(type(p_all[0].text))

def creatdataexcel(NEW_CONFIRMED_CASES_DATA_EXCEL,NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL,NEW_SPECIAL_ZONES_DATA_EXCEL):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('新增确诊人数')
    ws.write(0,1,'新增确诊人数')
    province = ['湖北','北京','天津','河北','山西','内蒙古','辽宁','吉林','黑龙江','上海','江苏','浙江','安徽','福建','江西','山东',
                '河南','湖南','广东','广西','海南','重庆','四川','贵州','云南','西藏','陕西','甘肃','青海','宁夏','新疆']
    column = 2
    for item in province:
        ws.write(0,column,item)
        column += 1
    wb.save(NEW_CONFIRMED_CASES_DATA_EXCEL)

    WB = xlwt.Workbook()
    WS = WB.add_sheet('新增无症状感染人数')
    WS.write(0,1,'新增无症状感染人数')
    province = ['湖北','北京','天津','河北','山西','内蒙古','辽宁','吉林','黑龙江','上海','江苏','浙江','安徽','福建','江西','山东',
                '河南','湖南','广东','广西','海南','重庆','四川','贵州','云南','西藏','陕西','甘肃','青海','宁夏','新疆']
    column = 2
    for item in province:
        WS.write(0, column, item)
        column += 1
    WB.save(NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL)

    WB2 = xlwt.Workbook()
    WS2 = WB2.add_sheet('港澳台累计确诊人数')
    province = ['台湾', '香港', '澳门']
    column = 1
    for item in province:
        WS2.write(0, column, item)
        column += 1
    WB2.save(NEW_SPECIAL_ZONES_DATA_EXCEL)



def new_confirmed_cases(data,row_count,blank_line,NEW_CONFIRMED_CASES_DATA_EXCEL):
    # print(data[0].text) # 疫情通报第一段
    try:
        rb = xlrd.open_workbook(NEW_CONFIRMED_CASES_DATA_EXCEL)  # 打开数据文件
        wb = copy(rb)  # 利用xlutils.copy下的copy函数复制
        ws = wb.get_sheet(0)  # 获取表单0
        sheet1 = rb.sheet_by_index(0)
        # 设置前两列宽度
        ws.col(0).width = 256 * 15
        ws.col(1).width = 256 * 15
        # 写入日期
        for item in data:
            months = re.findall('(\d+)月',item.text)
            days = re.findall('月(\d+)日',item.text)
            if len(months) & len(days) != 0:
                break
        ws.write(row_count,0,f'{int(months[0])}月{int(days[0])}日')
        print(f'{int(months[0])}月{int(days[0])}日')
        # 写入本土新增确诊人数
        # print(item.text) # 输出文本
        cases_counts = re.findall('本土病例(\d+)例', item.text)
        if len(cases_counts) == 0:
            cases_counts = re.findall('(\d+)例为本土病例', item.text)
            if len(cases_counts) == 0:
                cases_counts = re.findall('新增确诊病例(\d+)例', item.text)
                if len(cases_counts) == 0:
                    cases_counts = re.findall('我委收到国内\d+省（区、市）报告新增新型冠状病毒感染的肺炎确诊病例(\d+)例',item.text)
                    if len(cases_counts) == 0:
                        cases_counts = re.findall('新增病例(\d+)例', item.text)
                        if len(cases_counts) == 0:
                            if item.text.__contains__('均为境外输入病例'):
                                cases_counts.append(0)
                            else:
                                 return 0
        # print(item.text)
        # print(type(cases_counts)) # 类型为str
        # if int(cases_counts[0]) != 0:
        ws.write(row_count,1,int(cases_counts[0]))
        # 提取本土病例数据
        mainland_data = re.findall('本土病例\d+例（([\s\S]*?)）', item.text) # 要用正则的非贪婪模式，使得匹配最近的括号内容
        if len(mainland_data) == 0:
            mainland_data = re.findall('\d+例为本土病例（([\s\S]*?)）',item.text)
            if len(mainland_data) == 0:
                mainland_data = re.findall('报告新增确诊病例\d+例（([\s\S]*?)）',item.text)
                if len(mainland_data) == 0:
                    mainland_data = re.findall('报告，新增确诊病例\d+例（([\s\S]*?)）', item.text)
                    if len(mainland_data) == 0:
                        mainland_data = re.findall('我委收到国内\d+省（区、市）报告新增新型冠状病毒感染的肺炎确诊病例\d+例（([\s\S]*?)）',item.text)
                        if len(mainland_data) == 0:
                            mainland_data.append('无')

        province = ['湖北', '北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江', '上海', '江苏', '浙江',
                    '安徽', '福建', '江西', '山东',
                    '河南', '湖南', '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海',
                    '宁夏', '新疆']
        column = 2
        for arr in province:
            # print(mainland_data)
            infections1 = re.findall(f'{arr}(\d+)例',mainland_data[0])
            infections2 = re.findall(f'本土病例(\d+)例（*.在{arr}',item.text)
            infections3 = re.findall(f'其中(\d+)例为本土病例（均在{arr}）',item.text)
            infections4 = re.findall(f'含{arr}临床诊断病例(\d+)例',mainland_data[0])
            infections5 = re.findall(f'{arr}[省市](\d+)例',mainland_data[0])
            # print(infections1,end='')
            # print(infections2,end='')
            # print(infections3,end='')
            # print(infections4,end='')
            # print(infections5,end='')
            if len(infections1) != 0:
                ws.write(row_count, column, int(infections1[0]))
            elif len(infections2) != 0:
                ws.write(row_count, column, int(infections2[0]))
            elif len(infections3) != 0:
                ws.write(row_count, column, int(infections3[0]))
            elif len(infections4) != 0:
                ws.write(row_count, column, int(infections4[0]))
            elif len(infections5) != 0:
                ws.write(row_count, column, int(infections5[0]))
            else:
                ws.write(row_count,column,0)
            column += 1

        # # 特殊处理港澳台三地数据
        # for arr in data:
        #     infections1 = re.findall('台湾地区(\d+)例',arr.text)
        #     infections2 = re.findall('香港特别行政区(\d+)例',arr.text)
        #     infections3 = re.findall('澳门特别行政区(\d+)例',arr.text)
        #     if len(infections1) != 0:
        #         ws.write(page_count, column - 3, int(infections1[0]))
        #         info1 = sheet1.cell(page_count - 1, column - 3).value
        #         info12 = sheet1.cell(page_count - 2, column - 3).value
        #         if str(info1) != '台湾':
        #             ws.write(page_count - 1, column - 3,int(info1) - int(infections1[0]))
        #         elif str(info12) == '':
        #             ws.write(page_count - 2, column - 3, int(info12) - int(infections1[0]))
        #     if len(infections2) != 0:
        #         ws.write(page_count, column - 2, int(infections2[0]))
        #         info2 = sheet1.cell(page_count - 1, column - 2).value
        #         info22 = sheet1.cell(page_count - 2, column - 2).value
        #         if str(info2) != '香港':
        #             ws.write(page_count - 1, column - 2,int(info2) - int(infections2[0]))
        #         elif str(info22) == '':
        #             ws.write(page_count - 2, column - 2, int(info22) - int(infections1[0]))
        #     if len(infections3) != 0:
        #         ws.write(page_count, column - 1, int(infections3[0]))
        #         info3 = sheet1.cell(page_count - 1, column - 1).value
        #         info32 = sheet1.cell(page_count - 2, column - 1).value
        #         if str(info3) != '澳门':
        #             ws.write(page_count - 1, column - 1,int(info3) - int(infections3[0]))
        #         elif str(info32) == '':
        #             ws.write(page_count - 2, column - 1, int(info32) - int(infections1[0]))
        # if infections1 == 0:
        #     ws.write(page_count, column - 3 ,int(infections1[0]))
        # if infections2 == 0:
        #     ws.write(page_count, column - 2 ,int(infections2[0]))
        # if infections3 == 0:
        #     ws.write(page_count, column - 1,int(infections3[0]))
        wb.save(NEW_CONFIRMED_CASES_DATA_EXCEL)
    except Exception as e:
        print(f'新增确诊人数程序第{row_count}条出错')
        print(e)



def new_asymptomatic_infections(data,row_count,blank_line,NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL):
    try:
        rb = xlrd.open_workbook(NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL)  # 打开数据文件
        wb = copy(rb)  # 利用xlutils.copy下的copy函数复制
        ws = wb.get_sheet(0)  # 获取表单0
        sheet1 = rb.sheet_by_index(0)
        # 设置前两列宽度
        ws.col(0).width = 256 * 15
        ws.col(1).width = 256 * 20
        # 写入日期
        for item in data:
            months = re.findall('(\d+)月', item.text)
            days = re.findall('月(\d+)日', item.text)
            if len(months) & len(days) != 0:
                break
        ws.write(row_count, 0, f'{int(months[0])}月{int(days[0])}日')
        # print(f'{int(months[0])}月{int(days[0])}日')
        # 筛选符合要求的一段
        for item in data:
            if re.findall('31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者',item.text):
                break
        # 写入本土新增无症状感染者人数
        cases_counts = re.findall('新增无症状感染者\d+例，其中境外输入\d+例，本土(\d+)例', item.text)
        if len(cases_counts) == 0:
            cases_count1 = re.findall('新增无症状感染者(\d+)例（境外输入\d+例）', item.text)
            cases_count2 = re.findall('新增无症状感染者\d+例（境外输入(\d+)例）', item.text)
            if len(cases_count1) & len(cases_count2) != 0:
                cases_counts.append(int(cases_count1[0]) - int(cases_count2[0]))
            if len(cases_counts) == 0:
                if item.text.__contains__('（均为境外输入）'):
                    cases_counts.append(0)
                else:
                    blank_line.append(row_count)
                    return 0
        ws.write(row_count, 1, int(cases_counts[0]))
        # 提取本土病例数据
        mainland_data = re.findall('新增无症状感染者\d+例，其中境外输入\d+例，本土\d+例（([\s\S]*?)）', item.text)  # 要用正则的非贪婪模式，使得匹配最近的括号内容
        if len(mainland_data) == 0:
            mainland_data.append('无')
        province = ['湖北', '北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江', '上海', '江苏', '浙江',
                    '安徽', '福建', '江西', '山东',
                    '河南', '湖南', '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海',
                    '宁夏', '新疆']
        column = 2
        for arr in province:
            infections1 = re.findall(f'{arr}(\d+)例', mainland_data[0])
            infections2 = re.findall(f'本土(\d+)例（*.在{arr}）', item.text)
            if len(infections1) != 0:
                ws.write(row_count, column, int(infections1[0]))
            elif len(infections2) != 0:
                ws.write(row_count, column, int(infections2[0]))
            else:
                ws.write(row_count, column, 0)
            column += 1
        wb.save(NEW_ASYMPTOMATIC_INFECTIONS_DATA_EXCEL)
    except Exception as e:
        print(f'新增无症状感染者人数程序第{row_count}条出错')
        print(e)


def specialzones(data,row_count,NEW_SPECIAL_ZONES_DATA_EXCEL):
    try:
        rb = xlrd.open_workbook(NEW_SPECIAL_ZONES_DATA_EXCEL)  # 打开数据文件
        wb = copy(rb)  # 利用xlutils.copy下的copy函数复制
        ws = wb.get_sheet(0)  # 获取表单0
        for item in data:
            months = re.findall('(\d+)月', item.text)
            days = re.findall('月(\d+)日', item.text)
            if len(months) & len(days) != 0:
                break
        ws.write(row_count, 0, f'{int(months[0])}月{int(days[0])}日')
        sheet1 = rb.sheet_by_index(0)
        for arr in data:
            infections1 = re.findall('台湾地区(\d+)例', arr.text)
            infections2 = re.findall('香港特别行政区(\d+)例', arr.text)
            infections3 = re.findall('澳门特别行政区(\d+)例', arr.text)
            if len(infections1) != 0:
                ws.write(row_count, 1, int(infections1[0]))
                # info1 = sheet.cell(page_count - 1, 1).value
                # info12 = sheet.cell(page_count - 2, 1).value
                # if str(info1) != '台湾':
                #     ws.write(page_count - 1, 1, int(info1) - int(infections1[0]))
                # elif str(info12) == '':
                #     ws.write(page_count - 2, 1, int(info12) - int(infections1[0]))
                if len(infections2) != 0:
                    ws.write(row_count, 2, int(infections2[0]))
                # info2 = sheet.cell(page_count - 1, 2).value
                # info22 = sheet.cell(page_count - 2, 2).value
                # if str(info2) != '香港':
                #     ws.write(page_count - 1, 2, int(info2) - int(infections2[0]))
                # elif str(info22) == '':
                #     ws.write(page_count - 2, 2, int(info22) - int(infections1[0]))
            if len(infections3) != 0:
                ws.write(row_count, 3, int(infections3[0]))
                # info3 = sheet.cell(page_count - 1, 3).value
                # info32 = sheet.cell(page_count - 2, 3).value
                # if str(info3) != '澳门':
                #     ws.write(page_count - 1, 3, int(info3) - int(infections3[0]))
                # elif str(info32) == '':
                #     ws.write(page_count - 2, 3, int(info32) - int(infections1[0]))
            if infections1 == 0:
                ws.write(row_count, 1, int(infections1[0]))
            if infections2 == 0:
                ws.write(row_count, 2, int(infections2[0]))
            if infections3 == 0:
                ws.write(row_count, 3, int(infections3[0]))
        wb.save(NEW_SPECIAL_ZONES_DATA_EXCEL)
    except Exception as e:
        print(f'港澳台累计确诊人数程序第{row_count}条出错')
        print(e)
