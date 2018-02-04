#!/usr/bin/python
# -*- coding: utf-8 -*-
# @Time    : 2018/2/4 13:38
# @Author  : Derek.S
# @Site    : 
# @File    : graph.py

import xlrd
import networkx as nx

# open excel
def exceldata(filename, n):
    """
    open excel file read data
    :param filename: excel filename
    :return: None
    """
    try:
        book = xlrd.open_workbook(filename)
        sheets = book.sheets()
        sheet = sheets[0]
        dataset = []
        for r in range(sheet.nrows):
            col = sheet.cell(r, n).value
            if r != 0:
                dataset.append(col)
        return dataset
    except Exception as e:
        print(e)
        return None

# clean data
def cleandata(datalist, dictname, splitstat):
    """
    clean data
    :param datalist: datalist
    :return: dict file
    """
    with open(dictname, "w") as F:
        for datarow in datalist:
            if splitstat:
                for data in datarow.split(","):
                    if len(data) != 0:
                        F.writelines(data + "\n")
                    else:
                        pass
            else:
                F.writelines(datarow + "\n")
    F.close()



if __name__ == "__main__":
    cmp_list_raw = exceldata("graph.xlsx", 0)
    shareholder_list_raw = exceldata("graph.xlsx", 1)
    cleandata(cmp_list_raw, "cmpdict.txt", False)
    cleandata(shareholder_list_raw, "shareholder.txt", True)
