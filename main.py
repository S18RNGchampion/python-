import requests
import openpyxl
from bs4 import BeautifulSoup


# 定义函数 加载界面
def loadHtml(url):
    Details_H = requests.get(url, headers=headers)
    Details_H.encoding = "utf-8"
    Details = Details_H.text
    bf = BeautifulSoup(Details, "html.parser")
    return bf


# 定义函数 新建excel
def newExcel():
    wb = openpyxl.Workbook()  # 创建Excel对象
    ws = wb.active  # 获取当前正在操作的表对象
    ws.append(['票房排名', '电影名', '电影上映年份', '票房', '电影简介', '海报地址'])  # 设置excel分类
    return wb


# 爬取目标地址
url = 'http://58921.com'
# 添加请求信息与cookie
headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/79.0.3945.130 Safari/537.36",
    "Cookie": "DIDA642a4585eb3d6e32fdaa37b44468fb6c=aejslkiqdbs2mnbh28hsps9100;time"
              "=MTEzNTI2LjIxNjM0Mi4xMDI4MTYuMTA3MTAwLjExMTM4NC4yMDc3NzQuMTE5OTUyLjExMTM4NC4xMDQ5NTguMTE1NjY4LjEwOTI0Mi4xMTM1MjYuMTA0OTU4LjEwOTI0Mi4xMTU2NjguMTIyMDk0LjExOTk1Mi4xMDkyNDIuMA "
}
# 爬取网址
indexUrl = url + "/alltime"
# 读取首页
bf1 = loadHtml(indexUrl)
# 初始页码
NowPage = 0
# 总页码
TotalPage = bf1.find_all("span", "pager_number")[1].text.split('/')[1]

wb = newExcel()  #
wb19 = newExcel()  #
wb18 = newExcel()  #
wb17 = newExcel()  #
wb16 = newExcel()  # 均为准备工作
wb15 = newExcel()  #
num19 = 0  #
num18 = 0  #
num17 = 0  #
num16 = 0  #
num15 = 0  #
# while (NowPage <= int(TotalPage)):
while NowPage <= 3:
    for test in bf1.tbody.find_all('tr'):
        MovieName = test.find_all('td')[2].text  # 电影名
        MovieRank = test.find_all('td')[1].text  # 电影排行
        MovieTime = test.find_all('td')[6].text  # 上映时间
        MovieHref = url + test.a.get('href')  # 电影的地址
        bf2 = loadHtml(MovieHref)  # 加载电影地址
        bf3 = loadHtml(MovieHref + "/boxoffice")  # 电影票房地址
        MoviePrice = bf3.find_all("h3", "panel-title")[0].text.split()[1].split(')')[0]   # 电影票房
        MoviePaper = url + bf2.find_all('img')[1].get('src')  # 电影海报
        MovieIntro = bf2.find_all('div', 'panel-body content_view_content_body')[0].text  # 电影简介
        if MovieTime == '2019' and num19 <= 9:
            num19 = num19 + 1
            wb19.active.append([num19, MovieName, MovieTime, MoviePrice, MovieIntro, MoviePaper])  # 写入2019年
        if MovieTime == '2018' and num18 <= 9:
            num18 = num18 + 1
            wb18.active.append([num18, MovieName, MovieTime, MoviePrice, MovieIntro, MoviePaper])  # 写入2018年
        if MovieTime == '2017' and num17 <= 9:
            num17 = num17 + 1
            wb17.active.append([num17, MovieName, MovieTime, MoviePrice, MovieIntro, MoviePaper])  # 写入2017年
        if MovieTime == '2016' and num16 <= 9:
            num16 = num16 + 1
            wb16.active.append([num16, MovieName, MovieTime, MoviePrice, MovieIntro, MoviePaper])  # 写入2016年
        if MovieTime == '2015' and num15 <= 9:
            num15 = num15 + 1
            wb15.active.append([num15, MovieName, MovieTime, MoviePrice, MovieIntro, MoviePaper])  # 写入2015年
        wb.active.append([MovieRank, MovieName, MovieTime, MoviePrice, MovieIntro, MoviePaper])  # 添加进excel
        print(
            MovieRank + "\t" + MovieName + "\t" + MovieTime + "\t" + MoviePrice + "\t" + MovieIntro + "\t" + MoviePaper + "\t")
    NowPage = NowPage + 1  # 翻页
    bf1 = loadHtml(indexUrl + "?page=" + str(NowPage))  # 加载界面
# 存入所有信息
wb.save('main.xlsx')
wb19.save('19.xlsx')
wb18.save('18.xlsx')
wb17.save('17.xlsx')
wb16.save('16.xlsx')
wb15.save('15.xlsx')
