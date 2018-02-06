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
    with open(dictname, "w") as F:
        if splitstate:
            print("处理带分割符信息")
            



if __name__ == "__main__":
    exceldata("cxgs.xlsx", 1)