# -*- coding: utf-8 -*-
"""
Created on Sun Jan  3 13:27:57 2016

@author: YuKeji
"""
#匹配基因名与基因描述，合并各基因转录本的相同描述

#把字典中key中以某个字符串开始的key-value对都找出来
def match_keys_samestart(dict_to_find, word_to_search):
    match_data = {}
    for (key, value) in dict_to_find.items():
        if key.startswith(word_to_search):
            match_data[key] = value
    return match_data
    
#把字典中的value以列表的形式提取出来
def extract_values_aslist(dict_to_extract):
    a = []
    for (key, value) in dict_to_extract.items():
        a.append(dict_to_extract[key])
    return a
    
#去除list中的重复元素,并转换成字符串,以‘;’隔开
def merge_same_listelement(list_to_merge):
    new_list = list(set(list_to_merge))
    new_list2str = ';   '.join(new_list)#列表转str，用分号隔开
    return new_list2str
    

#将两个列表组成为键值对，构建字典
def creat_dic_from_2list(list1, list2):
    new_dict = dict(map(lambda x,y:[x,y],list1,list2))
    return new_dict
   
import xlrd
import xlwt
#读取被查找表格
book1 = xlrd.open_workbook('/Users/YuKeji/Desktop/TopBlast.xlsx')
sheet1 = book1.sheet_by_name('TopBlast')
gene_ID = sheet1.col_values(0)
gene_Description = sheet1.col_values(3)
#读取索引
book2 = xlrd.open_workbook('/Users/YuKeji/Desktop/new_file.xlsx')
sheet2 = book2.sheet_by_name('Sheet1')
list_to_find = sheet2.col_values(0)
#通过列表构建字典，引用函数
gene_dict = creat_dic_from_2list(gene_ID, gene_Description)
#新建excel文件，原文件只读
new_excel = xlwt.Workbook()
sheet_match = new_excel.add_sheet('sheet_match')

#初始行值为n = 0
n = 0
#历遍索引
for name in list_to_find:
    a = match_keys_samestart(gene_dict, name)#找出相同开头的键，构建新字典
    b = extract_values_aslist(a)#字典值转换为列表
    c = merge_same_listelement(b)#合并相同值，转换成str
    sheet_match.write(n, 2, c)#写入每个基因描述
    sheet_match.write(n, 0, name)#写入基因索引
    n = n+1 #迭代
new_excel.save('/Users/YuKeji/Desktop/ID_Match.xls')#存储为新文件






    
    