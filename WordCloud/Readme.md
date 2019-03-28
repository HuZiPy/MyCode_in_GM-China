install ***jieba*** and ***wordcloud*** base on the anaconda

  参考 [这个](https://blog.csdn.net/zhaohaibo_/article/details/79253740)
  
  大致步骤如下（windows）
    
  1. 在 [官网](https://pypi.org/project/wordcloud/) 下载 安装包
    
  2. 将安装包 解压至 anaconda 的 **pkgs** 目录下
    
  3. 通过终端（**Anaconda Prompt**）进入 安装包文件所在目录，输入命令  `python setup.py install` 
  
  
  
## StopWords.txt

该文件存的是 停顿词 数据


## userdict.txt

用户自定义字典，用于 ***jieba*** 分词


## excel.py  

输出词频

dictionary convert to DataFrame  

      pd.DataFrame(list(dict.items()), columns = ["a", "b"])
