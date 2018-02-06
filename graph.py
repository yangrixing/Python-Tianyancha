#!/usr/bin/python
# -*- coding: utf-8 -*-
# @Time    : 2018/2/4 13:38
# @Author  : Derek.S
# @Site    : 
# @File    : graph.py

import xlrd
import networkx as nx
import matplotlib.pyplot as plt
import codecs
import logging

logging.basicConfig(level=logging.INFO)


plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

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


def relation(filename):
    """
    processing relations
    :param filename: Excel filename
    :return: relations
    """
    relations = {}
    keys = []
    try:
        book = xlrd.open_workbook(filename)
        sheets = book.sheets()
        sheet = sheets[0]
        for r in range(sheet.nrows):
            cmp = sheet.cell(r, 0).value
            share = sheet.cell(r, 1).value
            shares = []
            for word in (share.split(",")):
                if word != "" and word != "暂无":
                    shares.append(word)
            if r != 0:
                relations.setdefault(cmp, shares)
        print("处理前：", len(relations))
        for key in relations.keys():
            if(len(relations[key]) == 0):
                keys.append(key)
        for empty_key in keys:
            relations.pop(empty_key)
        print("处理后：", len(relations))
        return relations
    except Exception as e:
        print(e)
        return None

# relations dict to tuples
def coverelation(relations):
    """
    cover relations dict to tuples
    :param relations: relations dict
    :return: tuples
    """
    relation_dicts = []
    for a in relations.keys():
        for word in relations[a]:
            relation_dicts.append((a, word))
    return relation_dicts


# clean data
def cleandata(datalist, dictname, splitstat):
    """
    clean xlsx data
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
def creategraph(picname, relation, dict1, dict2):
    colors = ['red', 'green', 'blue', 'yellow']
    cmpnode = []
    sharenode = []
    with open(dict1, "r", encoding="utf8") as f1:
        for node in f1.readlines():
            cmpnode.append(node.split("\n")[0])
    f1.close()
    with open(dict2, "r", encoding="utf8") as f2:
        for node in f2.readlines():
            sharenode.append(node.split("\n")[0])
    f2.close()
    DG = nx.DiGraph()
    print("节点较多，需要一定时间运行")
    DG.add_nodes_from(cmpnode)
    DG.add_nodes_from(sharenode)
    DG.add_edges_from(relation)
    #DG.add_nodes_from(['哈哈'])
    #DG.add_nodes_from(['呵呵'])
    #DG.add_edges_from([('哈哈', '呵呵')])
    nx.draw(DG, with_labels=True, node_size=300, font_size=8, node_color=colors)
    fig = plt.gcf()
    fig.set_size_inches(30, 30)
    fig.savefig(picname, dpi=600)


if __name__ == "__main__":
    excelfname = "graph.xlsx"
    cmp_list_raw = exceldata(excelfname, 0)
    shareholder_list_raw = exceldata(excelfname, 1)
    cleandata(cmp_list_raw, "cmpdict.txt", False)
    cleandata(shareholder_list_raw, "shareholder.txt", True)
    relations = relation(excelfname)
    # relation("graph.xlsx")
    relations_dict = coverelation(relations)
    creategraph("pic.png", relations_dict, "cmpdict.txt", "shareholder.txt")
