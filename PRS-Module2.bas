Attribute VB_Name = "Module2"
Sub CreateFileCopy()

Dim backupPath As String, destPathFull As String, fileName As String, folderName As String

'On Error GoTo errhandler

backupPath = "\\PSACLW02\ProjData\EnglandT\Misc\ProjectScorecards\"

If Right(backupPath, 1) <> "\" Then backupPath = backupPath & "\"

fileName = ThisWorkbook.Name

folderName = Dir(backupPath, vbDirectory)

If folderName = "" Then
    MkDir backupPath
    folderName = Dir(backupPath, vbDirectory)
End If

destPathFull = backupPath & Left(fileName, InStr(fileName, ".") - 1) & "-" & Format(Now, "yyyymmdd_hhnnss") & _
            Right(fileName, (Len(fileName) - InStr(fileName, ".") + 1))

If fileName > "" And folderName > "" Then
    ActiveWorkbook.SaveCopyAs destPathFull
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
