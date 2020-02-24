"""
面向过程程序
目的：
    爬取 url:'http://xkkm.sdzk.cn/zy-manager-web/html/xx.html'(全国各校的选考科目要求数据)
    将结果保存在Excel文件中
数据样式：
      学校   |  专业    |选科要求
    北京大学 |  英语    |不提科目要求
    北京大学 |  俄语    |不提科目要求
    北京大学 | 物理学类 |物理(1门科目考生必须选考方可报考)
"""

import requests
from lxml import etree
import xlwt

#创建Excel文档
xls = xlwt.Workbook()
sheet = xls.add_sheet('选科要求')
sheet.write(0, 0, '学校')
sheet.write(0, 1, '专业')
sheet.write(0, 2, '专业层次')
sheet.write(0, 3, '选科要求')
# 定义变量k，保存Excel数据时的行
k = 1

# 各个学校的dm,mc信息网址
url1 = "http://xkkm.sdzk.cn/zy-manager-web/html/xx.html"
# 各个学校选科要求的网址
url2 = "http://xkkm.sdzk.cn/zy-manager-web/gxxx/searchInfor"
# 头部伪装
head = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36',
    'Content-Type': 'text/html;charset=UTF-8'
}
response = requests.get(url=url1, headers=head)
# 编码处理（解决乱码问题）
response.encoding = None
# 把response.text规格化成HTML文件样式
html1 = etree.HTML(response.text)

# 利用xpath获取学校名和各学校的选考科目要求的网址、schools用于存放学校名称、dms用于存放学校的代号属性、mcs用于存放学校的名称属性
schools = html1.xpath('//div[@id="div5"]//tr/td[4]/text()')
dms = html1.xpath('//div[@id="div5"]//tr/td[5]/form/input[1]/@value')
mcs = html1.xpath('//div[@id="div5"]//tr/td[5]/form/input[2]/@value')

# 对每一个学校访问
for j in range(len(schools)):
    # 请求网页时的data数据
    data = {
        'dm': dms[j],
        'mc': mcs[j]
    }
    # 头部伪装
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.116 Safari/537.36'
    }
    response = requests.post(url=url2, data=data, headers=headers)
    html = etree.HTML(response.text)

    # xpath匹配专业和选科要求、levels用于存放专业层次（主要为区别本科专业和专科专业） pors用于存放各个学校的专业、limits用于存放各个专业的选科要求
    levels = html.xpath('//div[@id="ccc"]//tr/td[2]/text()')
    pros = html.xpath('//div[@id="ccc"]//tr/td[3]/text()')
    limits = html.xpath('//div[@id="ccc"]//tr/td[4]/text()')

    # 输出学校名称
    print('--------------------' + schools[j] + '-----------------')

    # for循环用于往Excel写入数据
    for i in range(len(pros)):
        # 数据清洗
        pros[i] = pros[i].replace('\r\n', '')
        pros[i] = pros[i].replace('\t', '')
        pros[i] = pros[i].replace(' ', '')
        limits[i] = limits[i].replace('\r\n', '')
        limits[i] = limits[i].replace('\t', '')
        limits[i] = limits[i].replace(' ', '')

        # 打印专业、和专业限制
        print(levels[i],  pros[i], limits[i])

        sheet.write(k, 0, schools[j])
        sheet.write(k, 1, pros[i])
        sheet.write(k, 2, levels[i])
        sheet.write(k, 3, limits[i])
        k += 1
# 保存数据到Excel中
xls.save('选科要求.xls')