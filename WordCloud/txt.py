# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 11:23:37 2019

@author: HuWei
"""

import jieba
# 加载 自定义字典
jieba.load_userdict("userdict.txt")
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import imageio

# 打开 需要 被分词的 txt文档
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
    wordlist = jieba.cut(text)
    wl = "/ ".join(wordlist)
    
    # 去除停用词
    ## 创建停用词 list
    f_stop_seg_list = open_txt("StopWords.txt").splitlines()

    AllWord = ''
    for myword in wl.split('/'):
        if not (myword.strip() in f_stop_seg_list) and len(myword.strip()) > 1:
            AllWord = AllWord + myword
            AllWord = AllWord + " "
            
    return AllWord


if __name__ == "__main__":
    
    # open the file 
    text = open_txt("ToCut.txt")
    
    # jieba processing text
    cutted_word = jieba_processing_txt(text)
    
    # Generate a word cloud image
    backing = imageio.imread('picture.jpg')
    
    wordcloud = WordCloud(background_color="white",mask=backing,max_words=2000,max_font_size=50,font_path='FangSong_GB2312.ttf').generate(cutted_word)
    
    # Display the generated image:
    # the matplotlib way:
    plt.figure(figsize=(15,10))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    plt.show()
        