Attribute VB_Name = "Module1"



Sub Ways_Update_TierData()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''  sheet S1-Insurance monthly PV sales'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''
'1. clear content in the S1-Insurance sheet

    Application.Workbooks("[working]Ways Price.xlsx").Worksheets("S1-Insurance monthly PV sales").Activate
    
    S1_lastRow = Range("b2").End(xlDown).Row
    S1_lastCol = 16
    
    Range("b2", Cells(S1_lastRow, S1_lastCol)).ClearContents


''''''''''''''''''''''''''''''''
'' 2. copy Insurance data from Tall.xlsx

    Dim wkb As Workbook
    
    path_ = "C:\Users\kzk0vj\Desktop\ways\Tall.xlsx"
    Set wkb = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)
    
    wkb.Sheets("sheet1").Activate
    Tall_lastRow = Range("a2").End(xlDown).Row
    Tall_lastCol = 15
    
    Range("a2", Cells(Tall_lastRow, Tall_lastCol)).Copy
    
    Application.Workbooks("[working]Ways Price.xlsx").Worksheets("S1-Insurance monthly PV sales").Activate
    
    Range("b2").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    
    ''''''''''''''autofill the A column's formula
    
    Range("a" & Range("a500").End(xlDown).Row).AutoFill Destination:=Range("a" & Range("a500").End(xlDown).Row, "a" & Tall_lastRow)
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''close Tall's Data
    ''''''if you want to check the data, you should delete the next code
        
        wkb.Close
    
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''  sheets S6-X Tier X Mix  ''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim arr5
arr5 = Array("S6-1 Tier 1 Mix", "S6-2 Tier 2 Mix", "S6-3 Tier 3 Mix", "S6-4 Tier 4 Mix", "S6-5 Tier 5 Mix")

For Each tier In arr5

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''1. clear content in the Tier X sheets'''''''''''''''''''''''

    Application.Workbooks("[working]Ways Price.xlsx").Worksheets(tier).Activate
    
    TX_lastRow = 0
    TX_lastCol = 0
    TX_lastRow = Range("b2").End(xlDown).Row
    TX_lastCol = 16
    
    Range("b2", Cells(S1_lastRow, S1_lastCol)).ClearContents
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' 2. copy Tier X data from S6-X Tier X Mix.xlsx in ways floder

    Dim TXwbk As Workbook
    
    path_ = "C:\Users\kzk0vj\Desktop\ways\" & tier & ".xlsx"
    Set TXwbk = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)
    
    TXwbk.Sheets("sheet1").Activate
    TX_lastRow = Range("a2").End(xlDown).Row
    TX_lastCol = 15
    
    Range("a2", Cells(TX_lastRow, TX_lastCol)).Copy
    
    Application.Workbooks("[working]Ways Price.xlsx").Worksheets(tier).Activate
    
    Range("b2").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    
    ''''''''''''''autofill the A column's formula
    
    Range("a" & Range("a500").End(xlDown).Row).AutoFill Destination:=Range("a" & Range("a500").End(xlDown).Row, "a" & TX_lastRow)
        
    TXwbk.Close

Next tier

End Sub

