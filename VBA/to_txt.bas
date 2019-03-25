Attribute VB_Name = "Module2"
Sub GenerateReport()
Attribute GenerateReport.VB_ProcData.VB_Invoke_Func = " \n14"
'
'

'
    Windows("Master File-Mar.xlsm").Activate
    Sheets("Daily_China Ldr").Activate

    days_ = Format(Now(), "d")
    colNum = Range("r23").Offset(0, days_ - 2).Column
    
    Range("r23", Cells(23, colNum)).Copy
    
    Windows("yoy-Retail.xlsm").Activate
    Sheets("2019").Activate
    
    ''''''
    'need update
    
    Range("C15").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    
    
    
    ''''''''''''''''''''''''
'UPDATE data
''''''''''''''

colNum1 = Format(Now() - 1, "d") + 2

Sheet9.Activate


'''''''''''''''''''''''
''' update every month

Sheet9.Range("b2") = Sheet1.Cells(15, colNum1)

Sheet9.Range("c2") = Sheet1.Cells(16, colNum1)

Sheet9.Range("e2") = Sheet1.Cells(17, colNum1)


'
' generate txt file
'
 
'Filename = Application.GetSaveAsFilename(fileFilter:="Text Files (*.txt), *.txt")


filename_ = "sgm_wuling_dly_" & Format(Now() - 1, "yyyymmdd") & "130000.txt"

'MsgBox filename_

filepath_ = "J:\Forecast-ST\Sales-Report\China Daily summary Report\EIS\SGM-Wuling EIS\2019\Mar\" & filename_
 
Open filepath_ For Output As #1

CC = ActiveSheet.UsedRange.Rows.Count
CB = ActiveSheet.UsedRange.Columns.Count
 
' a1

For i = 1 To 4

Print #1, Cells(1, i).Value & "|";  ' add ; to continue

Next

Print #1, Cells(1, 5).Value;

Print #1, Chr(10)

' a2
Print #1, Range("a2").Value & "|";

' b2
Print #1, Range("b2").Value & "|";

'c2
Print #1, Range("c2").Value & "|";

' d2
If Range("d2").Value > 0 Then
    Print #1, "+" & Format(Range("d2").Value, "0.0%") & "|";
Else
    Print #1, Format(Range("d2").Value, "0.0%") & "|";
End If

'e2
If Range("e2").Value > 0 Then

    Print #1, "+" & Format(Range("e2").Value, "0.0%")
Else
    
    Print #1, Format(Range("e2").Value, "0.0%")
End If

 
Close #1

End Sub
