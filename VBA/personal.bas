Attribute VB_Name = "Module1"
Sub deleteNoBrandName()
Attribute deleteNoBrandName.VB_Description = "delete no brand name and unhidden"
Attribute deleteNoBrandName.VB_ProcData.VB_Invoke_Func = "u\n14"


'''
' unhide
'''

Cells.EntireRow.Hidden = False

'
' use in excel :daily Sales Report
' delete the part of no brand name

For i = Range("c100").End(xlUp).Row To 6 Step -1

    If Range("c" & i) = "" Or Range("c" & i) = 0 Or Range("c" & i).Value = "#REF!" Then
        Range("c" & i).EntireRow.Delete
    End If

Next

Range("d6").Select

End Sub


Sub check_SGM_DailyData()

'''
' check SGM Daily Data Sum = Total ?   today
'''

For i = 6 To Range("c6").End(xlDown).Row

    ' check Buick Brand Total
    If Range("c" & i) = "Buick Brand Total" Then
    
        d = Format(Now() - 1, "d")
    
        bbt_col = d + 3  ' today report date column number
        
        bbt_row = i
        
        
        For j = 4 To bbt_col
        
            sum_ = 0
            For n = 6 To (bbt_row - 1)
            
                sum_ = sum_ + Cells(n, j)
                
            Next
        
            If sum_ <> Cells(bbt_row, j) Then
                
                MsgBox Cells(bbt_row, j).Value & " is error"
        
                Range(Cells(6, j), Cells(bbt_row - 1, j)).Interior.Color = 255
                'Cells(bbt_row, j).EntireRow.Select
            End If
        
        Next
        
    End If

    
    ' check Cadillac Brand Total
    If Range("c" & i) = "Cadillac Brand Total" Then
    
        d = Format(Now() - 1, "d")
    
        cbt_col = d + 3  ' today report date column number
        
        cbt_row = i
        
        
        For j = 4 To cbt_col
        
            sum_ = 0
            For n = bbt_row + 1 To (cbt_row - 1)
            
                sum_ = sum_ + Cells(n, j)
                
            Next
        
            If sum_ <> Cells(cbt_row, j) Then
                
                MsgBox Cells(cbt_row, j).Value & " is error"
        
                Range(Cells(bbt_row + 1, j), Cells(cbt_row - 1, j)).Interior.Color = 255
                'Cells(bbt_row, j).EntireRow.Select
            End If
        
        Next
        
    End If



     ' check Chevrolet Brand Total
    If Range("c" & i) = "Chevrolet Brand Total" Then
    
        d = Format(Now() - 1, "d")
    
        chbt_col = d + 3  ' today report date column number
        chbt_row = i
        
        
        For j = 4 To chbt_col
        
            sum_ = 0
            For n = (cbt_row + 1) To (chbt_row - 1)
            
                sum_ = sum_ + Cells(n, j)
                
            Next
        
            If sum_ <> Cells(chbt_row, j) Then
                
                MsgBox Cells(chbt_row, j).Value & " is error"
                
                Range(Cells(cbt_row + 1, j), Cells(chbt_row - 1, j)).Interior.Color = 255
                'Cells(chbt_row, j).Select
            End If
        
        Next
        
    End If
    
Next

MsgBox "F I N E"


End Sub


Sub EDW()

''''''''''''''''SGMW assm '''''''''''''''''''''''''
Workbooks("SGMWAssm.xlsx").Activate

lastRow1 = Range("a2").End(xlDown).Row
lastCol1 = 12

    Range("a2", Cells(lastRow1, lastCol1)).copy
    
    '''''''''' need update each month''''''''''''''''''
    Workbooks("China Daily Summary Report-production.xlsx").Sheets("Raw Data").Activate
    
    Range("C2").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False



'''''''''''''''' SGM  '''''''''''''''''''''''''
Workbooks("SGMProduction.xlsx").Sheets("Template").Activate

lastRow2 = Range("a2").End(xlDown).Row
lastCol2 = 13

    Range("a2", Cells(lastRow2, lastCol2)).copy
    
    '''''''''' need update each month''''''''''''''''''
    Workbooks("China Daily Summary Report-production.xlsx").Sheets("Raw Data").Activate
    
    Range("C126").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False


''''''''''''''''SGMW PT '''''''''''''''''''''''''

Workbooks("SGMWPT.xlsx").Activate

lastRow = Range("a2").End(xlDown).Row
lastCol = 12

    Range("a2", Cells(lastRow, lastCol)).copy
    
    '''''''''' need update each month''''''''''''''''''
    Workbooks("China Daily Summary Report-production.xlsx").Sheets("Raw Data").Activate
    
    Range("C622").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False



End Sub
