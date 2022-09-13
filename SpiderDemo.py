# 导入相关模块
import json
import time
import requests
import xlrd
import xlwt
from bs4 import BeautifulSoup
import re

# 发送请求，获取疫情首页内容
print('=====爬虫程序启动=====')
headers = {'User-Agent': 'Mozilla/5.0'}
response = requests.get('http://www.nhc.gov.cn/xcs/yqtb/list_gzbd.shtml',headers=headers)
while True:
    if response.status_code == 200:
        print("Main_page sucessed")
        break
    else:
        print("Main_page failed")
        time.sleep(1)
home_page = response.content.decode()
# print(home_page)

# 使用BeautifulSoup提取疫情数据
soup = BeautifulSoup(home_page,'lxml')
arr = soup.find_all('a') # 利用a标签提取各日疫情通报网站
print(arr)

# 导出url到excel当中
wb = xlwt.Workbook()
ws = wb.add_sheet("Url Sheet")  # 创建sheet
ws.write(0, 0, "国家卫健委各日疫情通报网站：")
count = 1  # 记录excel的行数
for term in arr:
    url = term['href'] # 提取href内容
    print(type(url))  # url类型为str
    print(url)
    print('http://www.nhc.gov.cn/'+url)
    while True:
        response = requests.get('http://www.nhc.gov.cn/'+url,headers=headers) # 获取每日疫情通报内容
        if response.status_code == 200:
            print("successed")
            break
        else:
            print("failed")
            time.sleep(1)
    data_page = response.content.decode()
    print(data_page)
    soup = BeautifulSoup(data_page, 'lxml')
    p_all = soup.find_all('p')
    print(p_all)

    ws.write(count,0,'http://www.nhc.gov.cn/'+url)
    count += 1
wb.save('C:/Users/vamos/Desktop/url_collections.xls')

# 读取excel到python当中
wb = xlrd.open_workbook('C:/Users/vamos/Desktop/url_collections.xls')
all_sheet_names = wb.sheet_names()
print(all_sheet_names)
# 根据索引获取sheet内容
sheet1_content = wb.sheet_by_index(0);
# 获取行当中的数据
rows = sheet1_content.row_values(1)
print(rows)
print(type(rows))  #row类型为list
# 将其转换为str类型，以便使用
url_str = rows[0]
print(url_str)
print(type(url_str))




# 获取标签中文本内容
for p in p_all:
    print(p.text)
    data_str = re.findall('([\s\S]*)',p.text)[0]
    data_str = re.findall('\[.+\]',p.text)
    print(data_str)

# python转换为json字符串
json_str = json.dumps(data_str,ensure_ascii=False)
print(json_str)

# python转换为json文件
with open('D:/python-code/data/test.json','a',encoding='UTF-8') as fp:
    json.dump(data_str,fp,ensure_ascii=False)
    corona_virus = json.loads(json_str)
    print(corona_virus)

print("=====爬虫程序完成=====")
