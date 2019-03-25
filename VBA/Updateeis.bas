Attribute VB_Name = "Module1"
Sub UpdateSummary()

'
' textToColumns
'

Sheet2.Activate
    Sheet2.Columns("A:A").Select
    Selection.textToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(50, 1), Array(62, 1), Array(81, 1), Array(98, 1), _
        Array(113, 1), Array(134, 1)), TrailingMinusNumbers:=True


' clear summary_sheet's content A-G
'' confirm the row and column number
Sheet1.Activate
    theLast_row = Sheet1.Range("a2").End(xlDown).Row
    theLast_column = 7

    ''choose the range and clear contents
    If Sheet1.Range("a2") <> "" Then
        Sheet1.Range("a2", Cells(theLast_row, theLast_column)).ClearContents
    End If

' copy data to goal path
' example
''''''''''''''''''''''''''''''''''''''''
'''Worksheets("Sheet1").Range("A1:D4").Copy _
'''    Destination:=Worksheets("Sheet2").Range("E5")

Sheet2.Activate

    the__Last_row = Sheet2.Range("a1").End(xlDown).Row
    the__Last_column = 7
    Sheet2.Range(Cells(1, 1), Cells(the__Last_row, the__Last_column)).Copy Sheet1.Range("a2")
    
' refress the PivotTable
Sheet1.Activate

    ThisWorkbook.RefreshAll


End Sub



Sub ClearConnents()

Sheet2.Cells.ClearContents

Range("a1").Select

End Sub


Sub To_YoY()

Sheet3.Activate

Sheet3.Range("C3").Select

End Sub


Sub Copy_SGM_Data()

'
' open the SGM Daily Report excel file
Dim wkb As Workbook

path_ = "C:\Users\kzk0vj\Desktop\SGM Daily Report.xlsx"
Set wkb = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)


''''''''''''''''''''''''''''''''''''''''''''''''
''wholesale
''''''''''''''''''''''''''''''''''''''''''''''''''''

wkb.Sheets("Wholesale").Activate
' copy SGM daily report data
lastRow = wkb.Sheets("Wholesale").Range("c6").End(xlDown).Row


lastCol = wkb.Sheets("Wholesale").Range("d" & lastRow).End(xlToRight).Column

' copy data
wkb.Sheets("Wholesale").Range("d" & lastRow, Cells(lastRow, lastCol)).Copy

'pastspecial paste:= xlpastevalues

'twkb.Sheets("Daily-Wholesale").Activate
ThisWorkbook.Sheets("YoY").Range("c3").PasteSpecial Paste:=xlPasteValues

Application.CutCopyMode = False


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'retail
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
wkb.Sheets("Retail").Activate

' copy SGM daily report data
lastRow1 = wkb.Sheets("Retail").Range("c6").End(xlDown).Row


lastCol1 = wkb.Sheets("Retail").Range("d" & lastRow1).End(xlToRight).Column

' copy data
wkb.Sheets("Retail").Range("d" & lastRow1, Cells(lastRow1, lastCol1)).Copy

'pastspecial paste:= xlpastevalues

'twkb.Sheets("Daily-Wholesale").Activate
ThisWorkbook.Sheets("YoY").Range("c9").PasteSpecial Paste:=xlPasteValues

Application.CutCopyMode = False


End Sub





