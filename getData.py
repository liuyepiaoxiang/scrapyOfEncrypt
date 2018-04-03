#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
import openpyxl
import datetime
from urllib import request
from bs4 import BeautifulSoup

FXH = 'https://www.feixiaohao.com/'
ALC = 'https://www.aicoin.net.cn/currencies'
BLC = 'https://block.cc/'
BLC_API = 'https://data.block.cc/api/v1'
coin = {}
FILENAME = '虚拟币爬取.xlsx'

## 获取时间
TODAY = datetime.date.today()

#新建excel
def creatwb(wbname):
    wb=openpyxl.Workbook()
    wb.save(filename=wbname)
    print ("新建Excel："+wbname+"成功")

def hasFile(wbname):
    wb = openpyxl.Workbook(wbname)
    sheet = wb.active
    if wb.active is None:
        creatwb(wbname)
    else:
        titleZh = ['序号', '代码', '名称', TODAY]
        # 判断表是否有数据
        if (sheet.max_row > 1 and sheet.max_column > 1):
            sheet.cell(row=1, column=sheet.max_column + 1, value=str(TODAY))
        else:
            for i in range(len(titleZh)):
                sheet.cell(row=1, column=i + 1, value=str(titleZh[i]))

def write07Excel(index, data):
    wb = openpyxl.load_workbook(FILENAME)
    sheet = wb.active
    sheet.title = 'sheet'

    # 从第二行开始写入数据
    for j in range(len(data)):
        sheet.cell(row=index + 2, column=j+1, value=str(list(data.values())[j]))

    wb.save(FILENAME)
    print("写入数据成功！")

if __name__ == "__main__":
    head = {}
    # 写入User Agent信息
    head['User-Agent'] = 'Mozilla/5.0 (Linux; Android 4.1.1; Nexus 7 Build/JRO03D) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.166  Safari/535.19'
    # 创建Request对象
    req = request.Request(FXH, headers=head)
    # 传入创建好的Request对象
    response = request.urlopen(req)
    # 读取响应信息并解码
    html = response.read().decode('utf-8')
    # 获取数据
    soup_texts = BeautifulSoup(html, 'lxml')
    texts = soup_texts.find(id='table')
    trs = texts.tbody.findAll('tr')
    # 判断表格是否存在
    hasFile(FILENAME)
    # 读取数据
    for index in range(len(trs)):
        tds = trs[index].findAll('td')
        coin.clear()
        coin.setdefault('index', tds[0].string)
        coin.setdefault('code', trs[index]['id'])
        coin.setdefault('name', tds[1].a.img['alt'])
        coin.setdefault('price', tds[3].a['data-cny'])
        #coin.setdefault('usPrice', tds[3].a['data-usd'])
        write07Excel(index, coin)



