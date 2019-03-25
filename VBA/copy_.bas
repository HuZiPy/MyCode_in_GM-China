Attribute VB_Name = "Module1"
Sub SGMWandSGM_Data()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''' UPDATE SGMW's data  paste '''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' open the SGMW Daily Report excel file

Dim sgmw_wkb As Workbook

path_ = "C:\!document\!daily report\C_Sinbox\SGMWSales.xlsx"
Set sgmw_wkb = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)


'''''1. wholesale

    sgmw_wkb.Sheets("Wholesale").Activate
    
    whs_lastCol = Format(Now(), "d") + 2
    whs_lastRow = Range("c6").End(xlDown).Row
    
    Range("d6", Cells(whs_lastRow, whs_lastCol)).Copy
    
    
    '''''''''' need update each month''''''''''''''''''
    Workbooks("Master File-Mar.xlsm").Sheets("Wholesale").Activate
    
    Range("g4").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    
    
'''''2. retail

    sgmw_wkb.Sheets("Retail").Activate
    
    r_lastCol = Format(Now(), "d") + 2
    r_lastRow = Range("c6").End(xlDown).Row
    
    Range("d6", Cells(r_lastRow, r_lastCol)).Copy
    
    
    '''''''''' need update each month''''''''''''''''''
    Workbooks("Master File-Mar.xlsm").Sheets("Retail").Activate
    
    Range("g4").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''' UPDATE SGM's data  paste '''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' open the SGM Daily Report excel file
    Dim wkb As Workbook
    
    path_ = "C:\Users\kzk0vj\Desktop\SGM Daily Report.xlsx"
    Set wkb = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''
'wholesale
'''''''''

    wkb.Sheets("Wholesale").Activate
    ' copy SGM daily report data
    lastRow = wkb.Sheets("Wholesale").Range("c6").End(xlDown).Row
  
    Dim x As Integer
    x = Format(Now() - 1, "d")
    
    lastCol = wkb.Sheets("Wholesale").Range("c6").Offset(0, x).Column
    
    ' copy data
    wkb.Sheets("Wholesale").Range("d6", Cells(lastRow, lastCol)).Copy
    
    
    '''''''''' need update each month''''''''''''''''''
    
    Workbooks("Master File-Mar.xlsm").Sheets("Wholesale").Activate
    Range("G37").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''
'retail
'''''''''

    wkb.Sheets("Retail").Activate
    ' copy SGM daily report data
    lastRow2 = wkb.Sheets("Retail").Range("c6").End(xlDown).Row
    
    lastCol2 = wkb.Sheets("Retail").Range("c6").Offset(0, x).Column
    
    ' copy data
    wkb.Sheets("Retail").Range("d6", Cells(lastRow2, lastCol2)).Copy
    
    'pastspecial paste:= xlpastevalues
    
    '''''''''' need update each month''''''''''''''''''
    Workbooks("Master File-Mar.xlsm").Sheets("Retail").Activate
    Range("G37").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False


End Sub

































