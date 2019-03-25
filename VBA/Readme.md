# VBA in Excel

## copy_.bas 

  转至 [这里](https://github.com/HuZiPy/MyCode_in_GM-China/blob/master/VBA/copy_.bas)
  
  知识点小计：
  
  1. 打开 excel 文件
  
    Dim sgmw_wkb As Workbook

    path_ = "filename"
    Set sgmw_wkb = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)

  
  2. 根据当天的日期选择性 复制数据 `Format(Now(), "d")`
  
  3. 粘贴类型（only value）
  
    Range("g4").PasteSpecial Paste:=xlPasteValues
    
  4. 基于某个 单元格 移动至另一个单元格，并返回 column number
    
    Range("c6").Offset(0, x).Column / Row
    
  5. 取消复制模式
    
    Application.CutCopyMode = False
    
    
    
## personal.bas

  转至 [这里](https://github.com/HuZiPy/MyCode_in_GM-China/blob/master/VBA/personal.bas)
  
  知识点小计：
  
  1. For 循环
  
    需要 **删除** 的时候 ： 倒序
      
      For i = Range("c100").End(xlUp).Row To 6 Step -1

        If Range("c" & i) = "" Or Range("c" & i) = 0 Or Range("c" & i).Value = "#REF!" Then
          Range("c" & i).EntireRow.Delete
        End If

      Next
      
    循环 数组 ( [ways](https://github.com/HuZiPy/MyCode_in_GM-China/blob/master/VBA/ways.bas) used )
    
      Dim arr5
      arr5 = Array("S6-1 Tier 1 Mix", "S6-2 Tier 2 Mix", "S6-3 Tier 3 Mix", "S6-4 Tier 4 Mix", "S6-5 Tier 5 Mix")

      For Each tier In arr5
      
        ''''''
      Next tier
  
  2. If 判断
  
      If ... Then
        
        ''''''
        
      End If
      
  3. 窗口显示
  
    MsgBox "hello world"
    
    
## ways.bas

  知识点小计：
  
  1. Array
  
    Dim arr5
    arr5 = Array("S6-1 Tier 1 Mix", "S6-2 Tier 2 Mix", "S6-3 Tier 3 Mix", "S6-4 Tier 4 Mix", "S6-5 Tier 5 Mix")
    
    For each ......
    
    
## eis.bas

excel 遇到个需求： 确定该值 是否 在 某区域内  => `countif()` formula

1. 另存为 没有 宏 的 excel 文件

      ThisWorkbook.SaveAs "filename", xlOpenXMLWorkbook
    
    
2. 删除控件

        For Each sht In Sheets
    
          sht.DrawingObjects.Delete
        
        Next

3. 发送邮件

## Updateeis.bas

excel formula `=DAY(TODAY()-1)&TEXT(TODAY()-1,"yyyy-mm-dd")&"31"`

  `TEXT()` 修改格式
  
  
## [to_txt.bas](https://github.com/HuZiPy/MyCode_in_GM-China/blob/master/VBA/to_txt.bas)

主要是 按照 需求（百分比、| etc.） 生成 txt 文件  部分


## modify_Type.bas

修改格式， 加总确定总数是否正确 etc.
