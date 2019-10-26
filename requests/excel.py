import os,sys
import requests
from bs4 import BeautifulSoup
from openpyxl import workbook  
from openpyxl import load_workbook

os.chdir('F:\\code\\python\\requests') #更改工作目录

wb = workbook.Workbook()  # 创建Excel对象
ws = wb.active  # 获取当前正在操作的表对象
# 往表中写入标题行,以列表形式写入！
ws.append(['序号', '日期', '姓名', '标题', '回复状态']) #写入标题   

n = 2#爬取页数
urllist = []#url存放列表
for i in range(1,n+1):
    url = "https://www1.szu.edu.cn/mailbox/list.asp?page={page}&leader=%CE%CA%CC%E2%CD%B6%CB%DF&tag=7".format(page = i)
    urllist.append(url)

subete = []    
time = []
title = []
sequence = []
name=[]
status = []

for url in urllist:          
    #URL = "https://www1.szu.edu.cn/mailbox/list.asp?leader=%CE%CA%CC%E2%CD%B6%CB%DF"
    r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64)'
                                             'AppleWebKit/537.36 (KHTML, like Gecko)'
                                             'Chrome/77.0.3865.90 Safari/537.36'})
    print(r.status_code)
    r.encoding = 'gbk'
    soup = BeautifulSoup(r.text, "lxml")
  
    #获取日期
    for k in soup.find_all("td",style="font-size: 9pt",width="85"):
        time.append(k.text.strip('\u3000'))
    
    
    #获取序号+姓名+状态
    for k in soup.find_all("td",align="center",style="font-size: 9pt"):
        subete.append(k.text)

    number = int(len(subete))#每页人数
    for j in range(0,number,3):
        sequence.append (subete[j])
        name.append(subete[j+1])
        status.append(subete [j+2])
   
    #获取标题
    for k in soup.find_all("a",class_="fontcolor3")[2:]:
        title.append(k.text.strip('·\xa0'))#去掉前缀·\xa0        
    print (title)

    #写入表格
    for x in range(20):
        ws.append([sequence[x], time[x], name[x], title[x], status[x]]) 
   
  
wb.save('mail.xlsx') 












