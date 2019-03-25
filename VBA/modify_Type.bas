Attribute VB_Name = "Module1"
Sub modify_()
'
' modify font, width, delete if all 0
'

'font & gridline

    ActiveWindow.DisplayGridlines = False
    With Cells.Font
        .Name = "Arial"
        .Size = 10
    End With
    Cells.ColumnWidth = 8.57
    Cells.Interior.Pattern = xlNone
        
        
For i = Range("a11111").End(xlUp).Row To 2 Step -1

    ' width
    If Range("a" & i).Value = "All New Cruze HB" Then
    
        Range("a" & i).Columns.AutoFit
        
    End If
    
    If Range("g" & i).Value = "All New Cruze HB" Then
        Range("g" & i).Columns.AutoFit
    End If
    
    If Range("a" & i).Value = "SGMW Chongqing" Then
        Range("a" & i).Columns.AutoFit
        
    End If
    
    ' delete if all 0
    If Range("a" & i) <> "" And Range("a" & i) <> "FAW-GM" Then
    
        If Range("b" & i) = 0 And Range("c" & i) = 0 And Range("d" & i) = 0 And Range("e" & i) = 0 And Range("h" & i) = 0 And Range("i" & i) = 0 And Range("j" & i) = 0 And Range("k" & i) = 0 And Range("a" & i).Value <> "SGMW Chongqing" Then
            Range("B" & i).EntireRow.Delete
        End If
    End If
    
    
Next
End Sub


Sub check_()

Dim sum_, n, i, j, startrow, endrow As Integer

For i = 2 To Range("a11111").End(xlUp).Row

    'check wholesale
    ' check Buick
    If Range("a" & i).Value = "Buick" Then
        
        'check Buick Day Data
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("a" & i).Offset(1, j).Row
            endrow = Range("a" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 1).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("a" & i).Offset(0, j).Value) >= 1 Then
                Range("a" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
    ' check Cadillac
    If Range("a" & i).Value = "Cadillac" Then
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("a" & i).Offset(1, j).Row
            endrow = Range("a" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 1).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("a" & i).Offset(0, j).Value) >= 1 Then
                Range("a" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
        
    ' check Chevy
    If Range("a" & i).Value = "Chevy" Then
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("a" & i).Offset(1, j).Row
            endrow = Range("a" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 1).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("a" & i).Offset(0, j).Value) >= 1 Then
                Range("a" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
    ' check Baojun
    If Range("a" & i).Value = "Baojun" Then
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("a" & i).Offset(1, j).Row
            endrow = Range("a" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 1).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("a" & i).Offset(0, j).Value) >= 1 Then
                Range("a" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
    
    ' check Wuling
    If Range("a" & i).Value = "Wuling" Then
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("a" & i).Offset(1, j).Row
            endrow = Range("a" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 1).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("a" & i).Offset(0, j).Value) >= 1 Then
                Range("a" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If

Next


'check retail
For i = 2 To Range("a11111").End(xlUp).Row

    ' check Buick
    If Range("G" & i).Value = "Buick" Then
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("G" & i).Offset(1, j).Row
            endrow = Range("G" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 7).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("G" & i).Offset(0, j).Value) >= 1 Then
                Range("G" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
    ' check Cadillac
    If Range("G" & i).Value = "Cadillac" Then
        
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("G" & i).Offset(1, j).Row
            endrow = Range("G" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 7).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("G" & i).Offset(0, j).Value) >= 1 Then
                Range("G" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
    
    ' check Cadillac
    If Range("G" & i).Value = "Cadillac" Then
        
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("G" & i).Offset(1, j).Row
            endrow = Range("G" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 7).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("G" & i).Offset(0, j).Value) >= 1 Then
                Range("G" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
    
    ' check Chevy
    If Range("G" & i).Value = "Chevy" Then
        
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("G" & i).Offset(1, j).Row
            endrow = Range("G" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 7).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("G" & i).Offset(0, j).Value) >= 1 Then
                Range("G" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
    
    ' check Baojun
    If Range("G" & i).Value = "Baojun" Then
        
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("G" & i).Offset(1, j).Row
            endrow = Range("G" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 7).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("G" & i).Offset(0, j).Value) >= 1 Then
                Range("G" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
    
    ' check Wuling
    If Range("G" & i).Value = "Wuling" Then
        
        
        For j = 1 To 4
        
            sum_ = 0
            
            startrow = Range("G" & i).Offset(1, j).Row
            endrow = Range("G" & i).Offset(1, j).End(xlDown).Row
            
            For n = startrow To endrow
            
                sum_ = sum_ + Cells(n, j + 7).Value
            Next
            
            'sum_ = Application.WorksheetFunction.Sum(Range("b" & startrow, "b" & endrow))
        
            If Abs(sum_ - Range("G" & i).Offset(0, j).Value) >= 1 Then
                Range("G" & i).Offset(0, j).Interior.Color = 255

                MsgBox " The part of red is error"
            End If
        Next
    End If
    
Next

MsgBox "no error"

End Sub


