Attribute VB_Name = "Module1"
Sub ModifyNameinEIS_WS()

''''''
' copy data
''''''

Dim wkb, twkb As Workbook

path_ = "C:\Users\kzk0vj\Desktop\SGM Daily Report.xlsx"
Set wkb = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

wkb.Sheets("Wholesale").Activate
' copy SGM daily report data
lastRow = wkb.Sheets("Wholesale").Range("c6").End(xlDown).Row

' the current month  have x days
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''need update
'''''''''''''''''''''
Dim x As Integer
x = 31

lastCol = wkb.Sheets("Wholesale").Range("c6").Offset(0, x + 1).Column

' copy data
wkb.Sheets("Wholesale").Range("c6", Cells(lastRow, lastCol)).Copy ThisWorkbook.Sheets("Daily-Wholesale").Range("b5")

'pastspecial paste:= xlpastevalues

'twkb.Sheets("Daily-Wholesale").Activate
' ThisWorkbook.Sheets("Daily-Wholesale").Range("b5").PasteSpecial Paste:=xlPasteValues

'''''''''''''''''''''''''''''''''''''''''''''''''''


ThisWorkbook.Sheets("Daily-Wholesale").Activate
Dim errorNum As Integer

lastRow = Range("b5").End(xlDown).Row

' copy B column to A column
Range("b5", "b" & lastRow).Copy
Range("a5").PasteSpecial xlPasteValues

' with the model update the name
'''  need update depend on the EIS sheet
For I = 5 To lastRow
    '''
    ' modify the name with EIS sheet
    '''
    
    If Range("a" & I).Value = "GL8 ES" Then
        Range("a" & I).Value = "GL8 ES 28T"
    End If
    
    If Range("a" & I).Value = "VELITE 5" Then
        Range("a" & I).Value = "VELITE"
    End If

    If Range("a" & I).Value = "New Malibu XL" Then
        Range("a" & I).Value = "Malibu XL"
    End If
    
    '''
Next


' fill formual in aj column
''''''''''' update each month  aj column
    '''''''''''''''''
Range("aj5").AutoFill Destination:=Range("aj5", "aj" & lastRow - 2)
    
For I = 5 To lastRow

    If Range("a" & I).Value = "Buick Brand Total" Or Range("a" & I).Value = "Cadillac Brand Total" Or Range("a" & I).Value = "Chevrolet Brand Total" Or Range("a" & I).Value = "Total" Then
    
        Range("a" & I).ClearContents
    Else
        If Range("aj" & I) <> 1 Then
        
            errorNum = errorNum + 1
            Range("a" & I).Interior.Color = 255
        End If
        
    End If
    
    
Next

If errorNum <> 0 Then
    MsgBox ("There is(are) " & errorNum & " error where fill red.")
End If

End Sub



Sub ModifyNameinEIS_RS()

''''''
' copy data
''''''

Dim wkb, twkb As Workbook

path_ = "C:\Users\kzk0vj\Desktop\SGM Daily Report.xlsx"
Set wkb = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)

'path_ = "C:\Users\kzk0vj\Desktop\test.xlsm"
'Set twkb = Application.Workbooks.Open(Filename:=path_, ReadOnly:=False)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

wkb.Sheets("Retail").Activate
' copy SGM daily report data
lastRow2 = wkb.Sheets("Retail").Range("c6").End(xlDown).Row

Dim x As Integer
x = 31         '''''   need update each month   '''''''''''

lastCol2 = wkb.Sheets("Retail").Range("c6").Offset(0, x + 1).Column

' copy data
wkb.Sheets("Retail").Activate
wkb.Sheets("Retail").Range("c6", Cells(lastRow2, lastCol2)).Copy ThisWorkbook.Sheets("Daily-Retail").Range("b5")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sheet1.Range("a2").Copy


ThisWorkbook.Sheets("Daily-Retail").Activate


''''''
' main solve different name between wholesale and retail
''''''

Dim errorNum As Integer

lastRow = Range("b5").End(xlDown).Row

' copy B column to A column
Range("b5", "b" & lastRow).Copy
Range("a5").PasteSpecial xlPasteValues

' with the model update the name
For I = 5 To lastRow
    '''
    ' modify the name with EIS sheet
    '''
    
    If Range("a" & I).Value = "GL8 ES" Then
        Range("a" & I).Value = "GL8 ES 28T"
    End If
    
    If Range("a" & I).Value = "VELITE 5" Then
        Range("a" & I).Value = "VELITE"
    End If

    If Range("a" & I).Value = "New Malibu XL" Then
        Range("a" & I).Value = "Malibu XL"
    End If
    
    '''
Next


' fill formual in AJ column
    
Range("aj5").AutoFill Destination:=Range("aj5", "aj" & lastRow - 2)
    
For I = 5 To lastRow

    If Range("a" & I).Value = "Buick Brand Total" Or Range("a" & I).Value = "Cadillac Brand Total" Or Range("a" & I).Value = "Chevrolet Brand Total" Or Range("a" & I).Value = "Total" Then
    
        Range("a" & I).ClearContents
    Else
        If Range("aj" & I) <> 1 Then
        
            errorNum = errorNum + 1
            Range("a" & I).Interior.Color = 255
        End If
        
    End If
    
    
Next

If errorNum <> 0 Then
    MsgBox ("There is(are) " & errorNum & " error where fill red.")
End If

End Sub


Sub check()

Dim err_ As Integer

For I = 3 To 244

    If Range("aj" & I) <> "" Then
        If Range("aj" & I) <> 0 Then
            err_ = err_ + 1
            
            Range("ag" & I).Interior.Color = 255
        End If
    End If
    
Next

If err_ <> 0 Then

    MsgBox ("There is(are) " & err_ & " error where fill red.")
Else
    MsgBox ("F I N E")
End If

End Sub

Sub save_as()

ThisWorkbook.Save

Application.DisplayAlerts = False

    'thisworkbook.SaveAs filename:= "J:\Forecast-ST\Sales-Report\China Daily summary Report\EIS\SGM EIS\2019\Jan\" & Format(Now() - 1, "yyyymmdd") & "_EIS report.xlsx"
    'ActiveWorkbook.SaveCopyAs Filename:="C:\tttt\" & Format(Now() - 1, "yyyymmdd") & "_EIS report.xlsx"
    
    ' clear marc code
    ThisWorkbook.SaveAs "J:\Forecast-ST\Sales-Report\China Daily summary Report\EIS\SGM EIS\2019\Mar\" & Format(Now() - 1, "yyyymmdd") & "_EIS report.xlsx", xlOpenXMLWorkbook
    
    '''
    ''''''''''''''''
    ''' paste value no formual
    'Sheets.Select
    'Cells.Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
     '   :=False, Transpose:=False
    'Application.CutCopyMode = False
    '''''''''''''''''''''''''''''''''''''''''''
    
    Dim sht As Worksheet
    
    ''''''''''
    '' clear active
    For Each sht In Sheets
    
        sht.DrawingObjects.Delete
        
        
        '''''''''''''''''''''
        'With sht
         '   For Each shp In .Shapes
          '      shp.Delete
           ' Next
        'End With
        ''''''''''''''''''''''''''
        
    Next
    
    Application.Sheets("email").Delete

    ActiveWorkbook.Save

    
    Application.DisplayAlerts = True
    
    
    
    path_ = "J:\Forecast-ST\Sales-Report\China Daily summary Report\EIS\SGM EIS\2019\Mar\!EIS report.xlsm"
    Application.Workbooks.Open Filename:=path_, ReadOnly:=False

End Sub


Sub send_mail()

Dim wkb_ As Workbook

Dim outlookapp As Outlook.Application
Set outlookapp = New Outlook.Application

Dim outlookitem As Outlook.MailItem
Set outlookitem = outlookapp.CreateItem(olMailItem)

to_ = Sheet6.Range("b1").Value
subject_ = Sheet6.Range("b2").Value
Content = Sheet6.Range("b3").Value
Attachment = Sheet6.Range("b4").Value
cc_ = Sheet6.Range("b5").Value


With outlookitem
    .To = to_
    .CC = cc_
    .Subject = subject_
    .Body = Content
    .Attachments.Add Attachment
    .Send
End With


End Sub


Sub tt()

Range("AI6") = Format(Now(), "mmm")

End Sub




