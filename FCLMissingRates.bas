Attribute VB_Name = "FCLMissingRates"
Option Explicit

Sub CreateReport()

Dim txt2columns As Range
Dim target As Range
Dim lr As Worksheet
Dim mr As Worksheet

Application.ScreenUpdating = False

Set lr = LatestReport
Set mr = MissingRates

lr.Columns(1).TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

'copying last report without last column
With lr.UsedRange
    .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count - 1).Copy
End With

'pasting it to first free row in Missing Rates
Set target = firstFree(mr, 1)
target.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

'disable auto filter in Missing Rates
On Error Resume Next
mr.ShowAllData
On Error GoTo 0

'text to columns in Missing Rates just to be sure ;)
mr.Columns(1).TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
mr.Range("A1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes

'fill all additional data and formats
Call fillData
Call fillFormats

MsgBox ("Added " & (firstFree(mr, 1).Row) - target.Row & " new lines.")

Application.ScreenUpdating = True

End Sub

Function firstFree(works As Worksheet, column As Long) As Range
Dim cell As Range

    For Each cell In works.Columns(column).Cells
        If cell.Value = "" Then
            Set firstFree = cell
            Exit For
        End If
    Next cell
End Function

Public Sub fillData()
Dim i As Long
Dim startRow As Long
Dim finishRow As Long
startRow = firstFree(MissingRates, 46).Row
finishRow = firstFree(MissingRates, 1).Row - 1
'column 46 - week no ; date ; test ; status ; date ; solved ; Comment ; Action taken ; Rootcause ; now with ; leadtime ; status2 leadtime range  last action add Missing rate

If Application.International(xlCountrySetting) = 48 Then
    For i = startRow To finishRow
        Cells(i, 45).Value = Format(Date, "yyyy-ww", vbMonday, vbFirstJan1)
        Cells(i, 46).Value = Format(Date, "yyyy-mm-dd", vbMonday, vbFirstJan1)
        Cells(i, 47).FormulaLocal = "=IFERROR(VLOOKUP(A" & i & ";'Latest Report'!A:A;1;0);0)"
        Cells(i, 48).FormulaLocal = "=IF(AU" & i & "=0; " & Chr(34) & "SOLVED" & Chr(34) & ";" & Chr(34) & "PENDING" & Chr(34) & ")"
        Cells(i, 54).FormulaLocal = "=(IF(AW" & i & ">0;NETWORKDAYS(AT" & i & ";AW" & i & ");NETWORKDAYS(AT" & i & ";TODAY())))-1"
        Cells(i, 55).FormulaLocal = "=IF(BB" & i & "<1;" & Chr(34) & "new" & Chr(34) & ";IF(AND(BB" & i & ">=1; BB" & i & "<6);" & Chr(34) & "pending" & Chr(34) & ";IF(BC" & i & ">=6; " & Chr(34) & "overdue" & Chr(34) & ")))"
    Next i
Else
    For i = startRow To finishRow
        Cells(i, 45).Value = Format(Date, "yyyy-ww", vbMonday, vbFirstJan1)
        Cells(i, 46).Value = Format(Date, "yyyy-mm-dd", vbMonday, vbFirstJan1)
        Cells(i, 47).FormulaLocal = "=IFERROR(VLOOKUP(A" & i & ",'Latest Report'!A:A,1,0),0)"
        Cells(i, 48).FormulaLocal = "=IF(AU" & i & "=0, " & Chr(34) & "SOLVED" & Chr(34) & "," & Chr(34) & "PENDING" & Chr(34) & ")"
        Cells(i, 54).FormulaLocal = "=(IF(AW" & i & ">0,NETWORKDAYS(AT" & i & ",AW" & i & "),NETWORKDAYS(AT" & i & ",TODAY())))-1"
        Cells(i, 55).FormulaLocal = "=IF(BB" & i & "<1," & Chr(34) & "new" & Chr(34) & ",IF(AND(BB" & i & ">=1, BB" & i & "<6)," & Chr(34) & "pending" & Chr(34) & ",IF(BB" & i & ">=6, " & Chr(34) & "overdue" & Chr(34) & ")))"
    Next i
End If


End Sub

Public Sub fillFormats()
Dim laRow As Long
laRow = firstFree(MissingRates, 1).Row - 1
MissingRates.Range(Cells(1, 1), Cells(laRow, 59)).Select

With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

MissingRates.Range("J1:J" & laRow & ", K1:K" & laRow & ", Z1:Z" & laRow & ", AV1:AW" & laRow & "").Select
Selection.Interior.ColorIndex = 15

End Sub

Public Sub Results()
Dim pend As Integer
pend = Application.WorksheetFunction.CountIfs(Range("AV:AV"), "PENDING", Range("BC:BC"), "pending")
Dim over As Integer
over = Application.WorksheetFunction.CountIfs(Range("AV:AV"), "PENDING", Range("BC:BC"), "overdue")
Dim news As Integer
news = Application.WorksheetFunction.CountIfs(Range("AV:AV"), "PENDING", Range("BC:BC"), "new")
Dim yesterday As Date
Dim solvedYesterday
yesterday = Application.WorksheetFunction.WorkDay(Date, -1)
MsgBox yesterday
solvedYesterday = Application.WorksheetFunction.CountIf(Range("AW:AW"), yesterday)

MsgBox "Today it is " & Date & ". There are: " & vbNewLine & news & " New Missing Rates" & vbNewLine & pend & " Pending missing rates" & vbNewLine & over & " Overdue missing rates" & vbNewLine & "Yesterday we solved " & solvedYesterday & " Cases"
End Sub