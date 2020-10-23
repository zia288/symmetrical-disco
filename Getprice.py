import requests
import os
import xlrd
from time import sleep
import re
import json
import time
import xlwt

w = xlwt.Workbook(encoding='utf-8')
data_sheet = w.add_sheet("price")
data_sheet.write(0, 0, 'ID')
data_sheet.write(0, 1, 'Num')
data_sheet.write(0, 2, '商品')
data_sheet.write(0, 3, '时间')
data_sheet.write(0, 4, '价格')
data_sheet.write(0, 5, '优惠券信息')
num = 1
ID = 1
for root_dir, sub_dir, files in os.walk("data"):     # 第一个为起始路径，第二个为起始路径下的文件夹，第三个是起始路径下的文件。
    for file in files:
        # 构造绝对路径
        file_name = os.path.join(root_dir, file)

        # 打开excel文件
        wb = xlrd.open_workbook(file_name)
        sh1 = wb.sheet_by_index(0)
        cols = sh1.col_values(4)
        cols_name = sh1.col_values(2)[1:]
        cols = cols[1:]
        print(cols)

        header = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36"
        }

        for i in range(len(cols)):
            param = {
                "url": cols[i]
            }
            r = requests.get("http://p.zwjhl.com/price.aspx", params=param, headers=header)
            res = r.text
            res = re.findall(r'flotChart.chartNow\(\'(.*)\',\'https:', res)[0]
            res = '['+res+']'
            res = json.loads(res)
            for data in res:
                timestamp = data[0] / 1000
                time_local = time.localtime(timestamp)
                dt = time.strftime("%Y-%m-%d %H:%M:%S", time_local)
                data_sheet.write(num, 0, ID)
                data_sheet.write(num, 1, num)
                data_sheet.write(num, 2, cols_name[i])
                data_sheet.write(num, 3, dt)
                data_sheet.write(num, 4, data[1])
                data_sheet.write(num, 5, data[2])
                num += 1
            sleep(1)
            ID += 1
            print(cols_name[i])
        sleep(2)
w.save('History_price.xls')