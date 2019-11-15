Attribute VB_Name = "Module1"
Sub NewWeek()

Dim WIP As Workbook, wipPath As String, wipDate As Date, lastDate As Date, lastRow As Integer
Dim wipSheet As Worksheet, i As Integer, j As Integer, k As Integer, x As Integer, coNum As Long, oldHrs As Single
Dim serialNum As String, engRate As Single, shipDate As String, wipLastRow As Integer, projExists As Boolean
Dim prodLine As String, machTyp As String, custName As String, newProjs() As String, testRate As Single

'''hardcoded values'''
wipPath = "\\PSACLW02\HOME\SHARED\CARR-CENTRITECH\Activity Logs\ScoreCards\"
engRate = 155
'''hardcoded values'''

On Error GoTo errhandler
Application.ScreenUpdating = False

'open WIP wkbk
For Each wB In Application.Workbooks
    If wB.Name Like "*WIP*.xls*" Then
        wB.Activate
        Set WIP = ActiveWorkbook
        Exit For
    End If
Next wB

If WIP Is Nothing Then
    On Error GoTo referr
    If Right(wipPath, 1) <> "\" Then
        wipPath = wipPath & "\"
    End If
    Set WIP = Workbooks.Open(Filename:=wipPath & "*WIP*.xls*", UpdateLinks:=True)
    On Error GoTo errhandler
End If

'check the last time this wkbk was updated
ThisWorkbook.Activate
Worksheets(1).Activate
lastDate = Range("A100").Value
Range("A99").Select
Selection.End(xlUp).Select
lastRow = ActiveCell.Row

'find date listed on WIP
WIP.Activate
For i = 1 To 12
    Range("A" & i).Select
    Selection.End(xlToRight).Select
    If ActiveCell.Column < 100 Then
        If InStr(1, ActiveCell.Value, "WEEK") > 0 Then
            wipDate = Right(ActiveCell.Value, Len(ActiveCell.Value) - InStrRev(ActiveCell.Value, " "))
            Exit For
        End If
    End If
Next i

If wipDate > 0 Then
'check if WIP date is right
    If lastDate >= wipDate Then
        Select Case MsgBox("There may be an issue with the WIP workbook. " & _
                            "It seems like it is outdated (from " & wipDate & _
                            ") and this workbook was last updated on " & lastDate & _
                            ". Would you like to cancel the update in order " & _
                            "to check on the WIP file?", vbYesNo, "WIP Issue")
        Case vbYes
            Exit Sub
        End Select
                            
    End If
End If

Application.ScreenUpdating = False
'copy new WIP projects to ENG wkbk
x = 0 'no new projects
For Each wipSheet In WIP.Sheets
    wipSheet.Activate
    
    'get last row
    Range("A100").Select
    Selection.End(xlUp).Select
    wipLastRow = ActiveCell.Row
    
    'get testing labor rate
    testRate = 0
    On Error Resume Next 'if an error occurs, dialog box will be used to fix it
        For i = 5 To 12
            For j = 1 To 20
                If InStr(1, Cells(i, j).Value, "@") > 0 Then
                    testRate = Right(Cells(i, j).Value, Len(Cells(i, j).Value) - InStr(Cells(i, j).Value, "@"))
                    Exit For
                End If
            Next j
        Next i
        
        testRate = testRate + 1
        If Err.Number > 0 Or testRate = 1 Then 'couldn't get it from the WIP
            testRate = InputBox("Unable to locate the labor rate for " & wipSheet.Name & " testing." & vbCrLf & _
                            vbCrLf & "Enter the labor rate for " & wipSheet.Name & " testing below. (units of $/hr)", _
                            "Testing Labor Rate")
        Else
            testRate = testRate - 1
        End If
    On Error GoTo errhandler
    
    'get co num / serial num and compare
    For i = 12 To wipLastRow
        wipSheet.Activate
        If Range("A" & i).Value > 0 Then
            projExists = False
            If IsNumeric(Range("E" & i).Value) Then 'CO num
                coNum = Range("E" & i).Value
                For j = 3 To lastRow
                    If ThisWorkbook.Worksheets(1).Range("D" & j).Value = coNum Then
                        projExists = True 'updateinfo
                        ThisWorkbook.Worksheets(1).Range("A" & j).Value = Range("A" & i).Value 'customer
                        ThisWorkbook.Worksheets(1).Range("B" & j).Value = Range("D" & i).Value 'type
                        ThisWorkbook.Worksheets(1).Range("C" & j).Value = wipSheet.Name 'prodline
                        ThisWorkbook.Worksheets(1).Range("E" & j).Value = Range("B" & i).Value 'ser num
                        If Range("G" & i).Value > 0 Then
                            ThisWorkbook.Worksheets(1).Range("K" & j).Value = Range("G" & i).Value 'conum
                        End If
                        Exit For
                    End If
                Next j
            End If
    
            If Not projExists Then 'check SN
                serialNum = Range("B" & i).Value
                For j = 3 To lastRow
                    If ThisWorkbook.Worksheets(1).Range("E" & j).Value = serialNum Then
                        projExists = True 'update info
                        ThisWorkbook.Worksheets(1).Range("A" & j).Value = Range("A" & i).Value 'customer
                        ThisWorkbook.Worksheets(1).Range("B" & j).Value = Range("D" & i).Value 'type
                        ThisWorkbook.Worksheets(1).Range("C" & j).Value = wipSheet.Name 'prodline
                        ThisWorkbook.Worksheets(1).Range("D" & j).Value = Range("E" & i).Value 'conum
                        If Range("G" & i).Value > 0 Then
                            ThisWorkbook.Worksheets(1).Range("K" & j).Value = Range("G" & i).Value 'conum
                        End If
                        Exit For
                    End If
                Next j
                
                If Not projExists Then
                    WIP.Activate
                    prodLine = wipSheet.Name
                    custName = Range("A" & i).Value
                    serialNum = Range("B" & i).Value
                    machTyp = Range("D" & i).Value
                    If IsNumeric(Range("E" & i).Value) Then
                        coNum = Range("E" & i).Value
                    Else
                        coNum = 0
                    End If
                    shipDate = Range("G" & i).Value
                    
                    'add project to list of new projects
                    ReDim Preserve newProjs(x)
                    newProjs(x) = UCase(prodLine) & " - " & UCase(custName)
                    x = x + 1
                    
                    'Add Row in proper spot
                    ThisWorkbook.Activate
                    k = 1
                    j = 0 'using to note when proper product line is reached
                    Do While k < 99
                        If UCase(Range("C" & k).Value) = UCase(prodLine) Then 'product line has been reached
                            j = j + 1
                        End If
                        
                        If UCase(Range("C" & k).Value) <> UCase(prodLine) Then 'insert row & add info to row
                            If j > 0 Then
                                Range(k & ":" & k).Insert CopyOrigin:=xlFormatFromRightOrBelow
                                Range("A" & k).Value = UCase(custName)
                                Range("B" & k).Value = UCase(machTyp)
                                Range("C" & k).Value = UCase(prodLine)
                                If coNum > 0 Then
                                    Range("D" & k).Value = coNum
                                End If
                                Range("E" & k).Value = UCase(serialNum)
                                Range("K" & k).Value = shipDate
                                Exit Do
                            End If
                        End If
                        k = k + 1
                    Loop
                    
                End If
            End If
        End If
    Next i
    WIP.Activate
Next wipSheet

ThisWorkbook.Activate
Worksheets(1).Activate

coNum = 0
'copy actual ENG hrs
For i = 3 To lastRow
    oldHrs = Range("O" & i).Value 'so that total hrs won't change after program runs
    Range("N" & i).Value = Range("O" & i).Value + Range("P" & i).Value
    Range("Q" & i).Value = Range("R" & i).Value + Range("S" & i).Value
    Range("X" & i).Value = Range("Y" & i).Value
    If IsNumeric(Range("D" & i).Value) And Range("D" & i).Value > 0 Then
        coNum = Range("D" & i).Value
    ElseIf IsNumeric(Right(Range("D" & i).Value, 6)) Then
        coNum = Right(Range("D" & i).Value, 6)
    Else
        serialNum = Range("E" & i).Value
    End If
    
    For j = 1 To WIP.Sheets.Count 'each sheet of WIP
        WIP.Activate
        WIP.Sheets(j).Activate
        Range("A100").Select
        Selection.End(xlUp).Select
        
        If coNum > 0 Or serialNum > "" Then
            For k = 10 To ActiveCell.Row 'each row in WIP sheet
                If coNum > 0 Then
                    If WIP.Sheets(j).Range("E" & k).Value = coNum Then
                        ThisWorkbook.Activate
                        Range("O" & i).Value = WIP.Sheets(j).Range("N" & k).Value / engRate
                        Range("R" & i).Value = WIP.Sheets(j).Range("P" & k).Value / testRate
                        Range("Y" & i).Value = WIP.Sheets(j).Range("K" & k).Value
                        Exit For
                    End If
                ElseIf serialNum > "" Then
                    If WIP.Sheets(j).Range("B" & k).Value = serialNum Then
                        ThisWorkbook.Activate
                        Range("O" & i).Value = WIP.Sheets(j).Range("N" & k).Value / engRate
                        Range("R" & i).Value = WIP.Sheets(j).Range("P" & k).Value / testRate
                        Range("Y" & i).Value = WIP.Sheets(j).Range("K" & k).Value
                        Exit For
                    End If
                End If
            Next k
        End If
    Next j
    
    'reconcile actual hours
    Range("P" & i).Value = Application.WorksheetFunction.Max(0, Range("P" & i).Value - (Range("O" & i).Value - oldHrs))
    
    coNum = 0
    serialNum = ""
    ThisWorkbook.Activate
Next i

WIP.Close savechanges:=False

'success msg
MsgBox "Numbers have been updated successfully"

If x > 0 Then
    MsgBox "New projects added:" & vbCrLf & vbCrLf & Join(newProjs, vbCrLf)
End If

Range("A100").Value = Date
Range("A3").Select
Range("A1").Select
Application.ScreenUpdating = True
Exit Sub

referr: MsgBox "Error opening the WIP reference workbook. Check that the directory:" & vbCrLf & vbCrLf & _
                wipPath & vbCrLf & vbCrLf & "contains an Excel file with 'WIP' in the name"
        Exit Sub

errhandler: On Error Resume Next
            WIP.Close savechanges:=False
            MsgBox "Unspecified error"
            Application.ScreenUpdating = True
            

End Sub

