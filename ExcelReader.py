# -*- coding: utf-8 -*-
"""
Created on Tue Aug 10 16:24:19 2021

@author: ASUS
"""

import xlrd


class ExcelReader():
    def __init__(self, filepath):
        self.filepath = filepath

    def read_excel(self):
        excelfile = xlrd.open_workbook(self.filepath)
        sheet = excelfile.sheet_by_index(0)
        row_num = sheet.nrows
        col_num = sheet.ncols
        # 读取excel信息并存成字典
        stu_info = []
        key = sheet.row_values(0)  # 用第一行数据作为字典的key值
        if row_num <= 1:  # 判断excel是否为空
            print("Excel表中没有任何学生数据")
        else:
            for i in range(1, row_num):
                each_info = {}
                info = sheet.row_values(i)
                for j in range(col_num):
                    # 把每一行的value赋值给改行的key
                    each_info[key[j]] = info[j]
                stu_info.append(each_info)
        return stu_info


if __name__ == '__main__':
    r = ExcelReader("./1801.xlsx")
    s = r.read_excel()
    print(s)
