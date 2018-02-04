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
    clen xlsx data
    :param datalist: data list
    :param dictname: dict txt name
    :param splitstat: bool cut or not
    :return: dict txt file
    """

    dataset_share = []
    dataset_share_uniq = []
    dataset_cmp_uniq = []
    with open(dictname, "w") as F:
        if splitstat:
            print("处理股东信息")
            for datarow in datalist:
                for data in datarow.split(","):
                    if len(data) != 0:
                        dataset_share.append(data + "\n")
            print("去重前：", len(dataset_share))
            dataset_share_uniq = sorted(set(dataset_share), key=dataset_share.index)
            print("去重后：", len(dataset_share_uniq))
            F.writelines(dataset_share_uniq)
        else:
            print("处理公司信息")
            dataset_cmp_uniq = sorted(set(datalist), key=datalist.index)
            print("去重前：", len(datalist))
            print("去重后：", len(dataset_cmp_uniq))
            for row in dataset_cmp_uniq:
                F.writelines(row + "\n")
        print("写入完毕，关闭文件")
        F.close()


# create network graph
def creategraph(picname, dict1, dict2):
    pass


if __name__ == "__main__":
    # cmp_list_raw = exceldata("graph.xlsx", 0)
    # shareholder_list_raw = exceldata("graph.xlsx", 1)
    # cleandata(cmp_list_raw, "cmpdict.txt", False)
    # cleandata(shareholder_list_raw, "shareholder.txt", True)
    pass