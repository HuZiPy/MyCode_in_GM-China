# Python for Data analysis

## [ways_tier.py](https://github.com/HuZiPy/MyCode_in_GM-China/blob/master/Python/ways_tier.py)

* 获得 py 文件 路径
	
		import os

		father_content = os.getcwd()
		
* [pd.pivot_table](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.pivot_table.html)

  聚合 这部分还得 再 琢磨 琢磨 `groupby`  etc   

  但感觉 excel 的 pivot table 功能也挺强大的  比如说 能得到 占 某类别 的百分比
	
  pandas 如何实现还不清楚
  

* [reset_index](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.reset_index.html)

后续 需要 注意的是 解决大小写的问题  例： "Van" 和 "VAN"

一个想法是 将原始数据 处理为全部 大（/小）写

	
## insurance.py

* 修改列名  `df.rename()`


* DataFrame 中插入 新的一列    `DOMESTIC_CV_df.insert(0,"Source","RETAIL")`


* DataFrame 合并 - 列方向    `concat`

		pd.concat([DOMESTIC_CV_df,DOMESTIC_PV_LOCAL_df,IMPORT_CV,IMPORT_PV],ignore_index=True)


## top.py

* 排序  `df.sort_values(by="CYTD",ascending=False)`
