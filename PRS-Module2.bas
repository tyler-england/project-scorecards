Attribute VB_Name = "Module2"
Sub CreateFileCopy()

Dim backupPath As String, destPathFull As String, fileName As String, folderName As String
Dim delFile As String, minDate As Date, oldestFile As String

On Error GoTo errhandler

backupPath = "\\PSACLW02\ProjData\EnglandT\Misc\Backups\Project_Scorecards\"

If Right(backupPath, 1) <> "\" Then
    backupPath = backupPath & "\"
End If

fileName = ThisWorkbook.Name

folderName = Dir(backupPath, vbDirectory)

If folderName = "" Then
    MkDir backupPath
    folderName = Dir(backupPath, vbDirectory)
End If

destPathFull = backupPath & Format(Now, "yyyymmdd_hhnnss") & "-" & fileName

If fileName > "" And folderName > "" Then
    ActiveWorkbook.SaveCopyAs destPathFull
    'delete oldest backup (to avoid having too many)
    fileName = Dir(backupPath & "*.xls*")
    minDate = FileDateTime(backupPath & fileName)
    oldestFile = fileName
    Do Until fileName = ""
        If FileDateTime(backupPath & fileName) < minDate Then
            oldestFile = fileName
            minDate = FileDateTime(backupPath & fileName)
        End If
        fileName = Dir()
    Loop
    If oldestFile <> "" Then
        On Error Resume Next 'cannot delete read-only files
            Kill (backupPath & oldestFile)
        On Error GoTo errhandler
    End If
    
End If

Exit Sub

errhandler:
If Application.UserName = "England, Tyler (PSA-CLW)" Then
    MsgBox "Unable to create database backup"
End If

End Sub

Sub UpdateAuxWkbks()

Dim wbPath As String, wbName As String, firstRow As Integer, i As Integer

Sheet1.Activate
Sheet1.Unprotect

'eng wkbk
firstRow = Application.WorksheetFunction.Match("Eng*", Range("A:A"), 0)
i = firstRow + 1
wbPath = Range("A" & i).Value

If Right(wbPath, 1) <> "\" Then
    wbPath = wbPath & "\"
End If

wbName = Dir(wbPath & "Eng*.xls*")

If wbName > "" Then
    
    i = firstRow + 3
    Range("A" & i).Value = Left(wbName, 10) & "..."
    
    i = firstRow + 4
    Range("A" & i).Formula = "=HYPERLINK(" & """" & wbPath & wbName & """" & _
                            "," & """" & "Engineering Manager Workbook" & """" & ")"
End If

'mfg wkbk
firstRow = Application.WorksheetFunction.Match("M*f*g*", Range("A:A"), 0)
i = firstRow + 1
wbPath = Range("A" & i).Value

If Right(wbPath, 1) <> "\" Then
    wbPath = wbPath & "\"
End If

wbName = Dir(wbPath & "*.xls*")

If wbName > "" Then

    i = firstRow + 3
    Range("A" & i).Value = Left(wbName, 10) & "..."
    
    i = firstRow + 4
    Range("A" & i).Formula = "=HYPERLINK(" & """" & wbPath & wbName & """" & _
                            "," & """" & "Manufacturing Workbook" & """" & ")"
                            
End If

Range("D1").Select

End Sub

Sub readWriteAccess()

'remove readonly
Dim readWrite As Boolean, i As Integer, rwNames() As String

i = 3
ReDim Preserve rwNames(i)
rwNames(0) = "ENGLAND"
rwNames(1) = "SHAW"
rwNames(2) = "VIOV"

readWrite = False
For Each rwName In rwNames
    'MsgBox Application.UserName & vbCrLf & InStr(UCase(Application.UserName), UCase(rwName))
    If InStr(UCase(Application.UserName), UCase(rwName)) > 0 Then
        readWrite = True
        Exit For
    End If
Next rwName
'MsgBox readWrite & ", " & ActiveWorkbook.ReadOnly

If readWrite Then
    If ActiveWorkbook.ReadOnly Then
        With ActiveWorkbook
            SetAttr .FullName, vbNormal
            .ChangeFileAccess xlReadWrite

            Application.DisplayAlerts = False
            .Save
            Application.DisplayAlerts = True
        End With
    End If
End If

End Sub

Sub changeTabs()

Dim wSheet1 As Worksheet, wSheet2 As Worksheet
Dim leftText As String, rightText As String, newName As String
Dim x As Boolean

For Each wSheet1 In ThisWorkbook.Sheets
    If wSheet1.Name <> "All Projects" Then
        If wSheet1.Name <> "A_New_Scorecard" Then
            If UCase(wSheet1.Range("D1").Value) <> "STOCK" Then
                leftText = Left(wSheet1.Name, InStr(wSheet1.Name, "-"))
                rightText = Left(UCase(wSheet1.Range("D1").Value), WorksheetFunction.Min(5, Len(wSheet1.Range("D1").Value)))
                newName = leftText & rightText
                x = True
                For Each wSheet2 In ThisWorkbook.Sheets
                    If wSheet2.Name = newName Then
                        MsgBox wSheet1.Name & " is in conflict because of " & wSheet2.Name
                        x = False
                    End If
                Next
                If x Then
                    wSheet1.Name = newName
                End If
            End If
        End If
    End If
Next


End Sub
