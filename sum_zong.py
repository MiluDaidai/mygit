#-*- coding: UTF-8 -*-
from openpyxl import load_workbook

class Doexcel:
    def __init__(self,filename,sheetname):
        self.filename=filename
        self.sheetname=sheetname

    def get_data(self):
        wb = load_workbook(self.filename)
        sheet = wb[self.sheetname]

        max_v=sheet.max_row
        sub_data = []

        for i in range(3,max_v+1):
            sub_data.append(sheet.cell(i,4).value)
        return sub_data

    def get_name(self):
        wb = load_workbook(self.filename)
        sheet = wb[self.sheetname]

        max_v=sheet.max_row
        sub_name = []

        for i in range(3,max_v+1):
            sub_name.append(sheet.cell(i,2).value)
        return sub_name

import requests
class HttpRequests:
    def HttpRequest(self, url, method,cookie):

        requests.packages.urllib3.disable_warnings()
        try:
            if method.lower()=='get':
                res = requests.get(url, cookies = cookie, verify=False) 
            else:
                res = requests.post(url, cookies = cookie ,verify=False) 

            if res.status_code == 200:
                print(url + "访问成功")
            else :
                print(url + "访问失败")

            return res
        except Exception as e:
            print(url + "访问失败")
            # print(e)

            pass

if __name__ == '__main__':

    url = Doexcel("D:\product\公司自有产品培训演示环境及培训计划表20220627.xlsx", "Sheet1").get_data()
    name = Doexcel("D:\product\公司自有产品培训演示环境及培训计划表20220627.xlsx", "Sheet1").get_name()

    for i in range(0, len(url)):
        url1 = url[i]
        print 'name[i],end=' ''
        result = HttpRequests().HttpRequest(url=url1, method='get', cookie=None)
