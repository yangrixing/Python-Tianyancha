#!/usr/bin/python
# -*- coding: utf-8 -*-
# @Time    : 2018/2/5 15:56
# @Author  : Derek.S
# @Site    : 
# @File    : analyze.py

import xlrd

def exceldata(filename, n):
    """
    open excel file read data
    :param filename:  excel file name
    :param n: sheet number
    :return:  xlrd sheet data
    """
    try:
        book = xlrd.open_workbook(filename)
        sheets = book.sheets()
        sheet = sheets[n]
        return sheet
    except Exception as e:
        print(e)
        return None


def cleandata(sheet, datalist, dictname, splitstate):
    """
    clean xlsx data and create dict
    :param sheet: excel sheet
    :param datalist: data row
    :param dictname: dict txt file name
    :param splitstate: rows string splite switch
    :return: None
    """
    data = []
    data_uniq = []
    with open(dictname, "w") as F:
        if splitstate:
            print("处理带分割符信息")
            for datarow in range(sheet.nrows):
                # 处理表头
                if datarow != 0:
                    col = sheet.cell(datarow, datalist).value.split(",")
                    for word in col:
                        if word != "" and word != "暂无":
                            data.append(word)
            data_uniq = sorted(set(data), key=data.index)
            print("信息去重")
            for word in data_uniq:
                F.writelines(word + "\n")
    F.close()



if __name__ == "__main__":
    sheet = exceldata("result.xlsx", 0)
    cleandata(sheet, 16, "share_dict.txt", True)