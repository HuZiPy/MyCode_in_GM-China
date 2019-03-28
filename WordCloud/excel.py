# -*- coding: utf-8 -*-
"""
Created on Thu Mar 28 18:57:05 2019

@author: HuWei
"""


import pandas as pd
import collections
import jieba
# 加载 自定义字典
jieba.load_userdict("userdict.txt")


# 获取 excel 内的comment
def open_excel(filename):
    '''
    打开需要被分词的 excel 文件
    '''
    df = pd.read_excel(filename)
    
    text = df["comment"].sum()
    
    return text
    
def open_txt(filename):
    '''
    打开需要被分分词的 文本 文件  txt -  utf-8 格式
    '''
    with open(filename, encoding='utf-8') as f:
        text = f.read()
    
    return text


def jieba_processing_txt(text_):
    '''
    1. 分词    2. 去除停用词
    '''
    # 分词
    wordlist = jieba.cut(text_)
    wl = "/ ".join(wordlist)
    
    # 去除停用词
    ## 创建停用词 list
    f_stop_seg_list = open_txt("StopWords.txt").splitlines()

    AllWord = []
    for myword in wl.split('/'):
        if not (myword.strip() in f_stop_seg_list) and len(myword.strip()) > 1:
            AllWord.append(myword)
            
    return AllWord


if __name__ == "__main__":
    
    # open the file 
    text = open_excel("comments_Data.xlsx")
    
    # jieba processing text
    cutted_word = jieba_processing_txt(text)
    
    # counter the number
    dic = dict( collections.Counter(cutted_word) )
    
    # dict covert to DataFrame
    arr = list(dic.items())
    
    frequence_df = pd.DataFrame(arr, columns = ["world", "frequence"])
    
    # 排序 高-> 低
    frequence_df = frequence_df.sort_values(by="frequence",ascending=False)
    
    frequence_df.to_excel("frequence.xlsx", index = False)