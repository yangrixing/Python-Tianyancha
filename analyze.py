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
        print(sheet)
        return sheet
    except Exception as e:
        print(e)
        return None


if __name__ == "__main__":
    exceldata("cxgs.xlsx", 0)