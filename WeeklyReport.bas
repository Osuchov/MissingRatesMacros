Attribute VB_Name = "WeeklyReport"
Option Explicit

Sub GoGoReport()

Dim directory As String
Dim mailTitles() As Variant
Dim mail As Variant
Dim mailTitleString As String
Dim test As Boolean

directory = pickDir("Pick the directory with Missing rate reports e-mails!", "Go Go Report!")

If Len(directory) = 0 Then
    MsgBox "Directory not picked. Exiting...", vbExclamation
    Exit Sub
End If

'Application.ScreenUpdating = False
mailTitles = Array("Missing FCL Rates.msg", "Missing LCL Rates.msg", "Missing Rates AIR.msg", "Missing Rates Road Europe -NL28 IT59 GB71 NL59 RO59.msg")
test = DirTest(directory, mailTitles)   'test if the picked dir contains valid missing rate files

If test = False Then    'MsgBox if DirTest failed
    mailTitleString = ""
    For Each mail In mailTitles
        mailTitleString = mailTitleString & mail & vbCrLf
    Next mail
    
    MsgBox "Picked directory does not contain valid files (4 missing rate e-mails):" & vbCrLf _
            & mailTitleString
    GoTo Finish
End If

Finish:
'    Application.CutCopyMode = False
'    wb.Close False
'    file = Dir

Application.ScreenUpdating = True

End Sub


Function pickDir(winTitle As String, buttonTitle As String) As String

Dim window As FileDialog
Dim picked As String

Set window = Application.FileDialog(msoFileDialogFolderPicker)
window.Title = winTitle
window.ButtonName = buttonTitle

If window.Show = -1 Then
    picked = window.SelectedItems(1)
    If Right(picked, 1) <> "\" Then
        pickDir = picked & "\"
    Else
        pickDir = picked
    End If
    
End If

End Function

Function DirTest(directory As String, mailTitles As Variant) As Boolean
    Dim file As String
    Dim allFiles As Long
    
    allFiles = 0
    file = Dir(directory & "*.msg")
    
    Do While file <> ""
        allFiles = allFiles + 1
        If IsInArray(file, mailTitles) = False Then
            DirTest = False
            Exit Function
        End If
        file = Dir()
    Loop
    
    If allFiles = 4 Then
        DirTest = True
    Else
        DirTest = False
    End If
    
End Function

Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise
Dim element As Variant

On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function

IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function
