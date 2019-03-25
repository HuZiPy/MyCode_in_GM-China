# Python for Data analysis

## [ways_tier.py](https://github.com/HuZiPy/MyCode_in_GM-China/blob/master/Python/ways_tier.py)

* 获得 py 文件 路径
	
		import os

		father_content = os.getcwd()
		
* [pd.pivot_table](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.pivot_table.html)

  聚合 这部分还得 再 琢磨 琢磨 `groupby`  etc   

  但感觉 excel 的 pivot table 功能也挺强大的  比如说 能得到 占 某种类 的百分比
	
  pandas 如何实现还不清楚
  

* [reset_index](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.reset_index.html)

后续 需要 注意的是 解决大小写的问题  例： "Van" 和 "VAN"

一个想法是 将原始数据 处理为全部 大（/小）写

	
