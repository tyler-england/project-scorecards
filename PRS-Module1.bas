Attribute VB_Name = "Module1"
Sub test()

For Each ws In ThisWorkbook.Sheets
    ws.Activate
    ws.Unprotect
    Range("F21").HorizontalAlignment = xlCenter
    Range("F21").WrapText = False
    ws.Protect
Next ws
Exit Sub

Dim rwNames(2) As String, i As Integer

rwNames(0) = "Zero"
rwNames(1) = "one"

For i = 0 To 1
    MsgBox rwNames(i)
Next i


End Sub
Sub AddNewProjects()

Dim wipPath As String, WIP As Workbook, prodLine As String, i As Integer, j As Integer, k As Integer, wB As Workbook
Dim orderNum As Single, rowLimit As Integer, numCheck As Integer, lastRow As Integer, numCols As Integer, booContinue As Boolean
Dim customeR As String, tyP As String, sN As String, montH As String, sellPrice As Single, lastDate As Date, costText As String
Dim matlCost As Single, mfgCost As Single, engCost As Single, evoCost As Single, newDataPoint As Boolean, testCost As Single
Dim priorCO As Single, colLet As String, wipDate As Date, updateType As String, snLong As String, newOrder As Boolean
Dim engPath As String, engFile As String, mfgPath As String, mfgFile As String, engWB As Workbook, mfgWB As Workbook, marRow As Integer

On Error GoTo errhandler

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
counter = 1
Call CreateFileCopy 'back up file
ThisWorkbook.Activate
Sheet1.Activate
wipPath = Range("A2").Value
If Right(wipPath, 1) <> "\" Then
    wipPath = wipPath & "\"
    'wipPath ends with \
End If

indexBurt = Application.WorksheetFunction.Match("Burt", Range("A:A"), 0)

newDataPoint = False
Select Case MsgBox("Do you want to update each project scorecard?" & vbCrLf & vbCrLf & _
            "Clicking 'yes' will irreversibly update scorecard data. Clicking 'no' will " & _
            "only update the info on the 'All Projects' sheet.", vbYesNoCancel, "Update")

Case vbYes
    newDataPoint = True
Case vbCancel
    Exit Sub
End Select

priorCO = 0
'for initial CO number
Call UpdateAuxWkbks 'set PM wkbk & MFG wkbk
engPath = Range("A" & Application.WorksheetFunction.Match("Eng*", Range("A:A"), 0) + 1).Value
If Right(engPath, 1) <> "\" Then
    engPath = engPath & "\"
End If
engFile = Left(Range("A" & Application.WorksheetFunction.Match("Eng*", Range("A:A"), 0) + 3).Value, 5)
mfgPath = Range("A" & Application.WorksheetFunction.Match("M*f*g*", Range("A:A"), 0) + 1).Value
If Right(mfgPath, 1) <> "\" Then
    mfgPath = mfgPath & "\"
End If
mfgFile = Left(Range("A" & Application.WorksheetFunction.Match("M*f*g*", Range("A:A"), 0) + 3).Value, 5)
'open WIP wkbk
For Each wB In Application.Workbooks
    If wB.Name Like "*WIP*.xls*" Then
        wB.Activate
        Set WIP = ActiveWorkbook
    End If
Next wB

On Error GoTo referr
If WIP Is Nothing Then
    Set WIP = Workbooks.Open(fileName:=wipPath & "*WIP*.xls*", UpdateLinks:=True)
End If
On Error GoTo errhandler

'find date listed on WIP
wipDate = 0
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
    If Date - wipDate > Weekday(Date) Then
        Select Case MsgBox("There may be an issue with the WIP workbook. " & _
                            "It seems like it is outdated (from " & wipDate & _
                            "). Would you like to cancel the update in order " & _
                            "to check on the WIP file?", vbYesNo, "WIP Issue")
        Case vbYes
            engWB.Close savechanges:=False
            mfgWB.Close savechanges:=False
            Exit Sub
        End Select
                            
    End If
End If

'add projects
For i = indexBurt To indexBurt + 2

    ThisWorkbook.Activate
    
    prodLine = Range("A" & i).Value

    WIP.Activate
    WIP.Worksheets(prodLine).Activate
    Range("A1000").Select 'go below machine list
    Selection.End(xlUp).Select 'go up to latest entry
    rowLimit = Application.WorksheetFunction.Max(ActiveCell.Row, 13)

    For j = 12 To rowLimit 'go through rows of WIP
    
        newOrder = False
        WIP.Activate
        
        If Range("A" & j).Value > 0 Then
            'machine exists
            newOrder = True
            
            snLong = Cells(j, Application.WorksheetFunction.Match("*SERIAL*", Range("11:11"), 0)).Value
            If snLong = "" Then
                snLong = "NONE-" & Str(Round(Rnd(), 3))
            End If
            
            'set sN
            If Left(snLong, 1) = "W" Then
            sN = Right(snLong, Len(snLong) - InStr(snLong, "-"))
            ElseIf Left(snLong, 1) = "C" And Len(snLong) < 10 Then
                sN = Right(Left(snLong, 5), 4)
            Else
                sN = snLong
            End If
            
            'get all the info for that row
            customeR = Range("A" & j).Value
            tyP = Cells(j, Application.WorksheetFunction.Match("*TYPE*", Range("11:11"), 0)).Value
            
            On Error Resume Next
            orderNum = Cells(j, Application.WorksheetFunction.Match("*ORDER*", Range("11:11"), 0)).Value
            If orderNum = 431767 Then 'one-off for cotter proj may 2020
                orderNum = 0
                sN = Right(snLong, Len(snLong) - 1)
            End If
            If orderNum = priorCO Then
                'when order number is a string, the macro fails to read
                'it and keeps the last value of orderNum
                orderNum = 0
            End If
            priorCO = orderNum
            
            montH = Cells(j, Application.WorksheetFunction.Match("*SHIPPED*", Range("11:11"), 0)).Value
            sellPrice = Cells(j, Application.WorksheetFunction.Match("*PRICE*", Range("11:11"), 0)).Value
            matlCost = Cells(j, Application.WorksheetFunction.Match("*MATERIAL*", Range("10:10"), 0)).Value
            mfgCost = Cells(j, Application.WorksheetFunction.Match("*MANUFACTUR*", Range("10:10"), 0)).Value
            engCost = Cells(j, Application.WorksheetFunction.Match("*ENGINEER*", Range("10:10"), 0)).Value
            evoCost = Cells(j, Application.WorksheetFunction.Match("*ROYAL*", Range("11:11"), 0)).Value
            testCost = Cells(j, Application.WorksheetFunction.Match("*TEST*", Range("11:11"), 0)).Value
            
            If tyP = "" Then
                tyP = "NONE"
            End If
            If montH = "" Then
                montH = "NONE"
            End If
            
            On Error GoTo errhandler
            
            'see if it needs to be added
            ThisWorkbook.Activate
            Sheet1.Activate
            
            For k = 1 To 100
                If Cells(5, k).Value = snLong Then
                    newOrder = False
                    Exit For
                End If
            Next k
            
            If newOrder Then
                
                'put everything into this wkbk
                ThisWorkbook.Activate
                
                'check number of current machines for that product line
                numCheck = Application.WorksheetFunction.Match(prodLine, Range("1:1"), 0)
                
                If Cells(1, numCheck + 1).Value = 0 And Cells(2, numCheck + 1).Value > 0 Then
                    'next machine is in same product line
                    'good to go, just insert a column
                    Columns(numCheck + 1).Insert shift:=xlToRight, _
                        CopyOrigin:=xlFormatFromLeftOrAbove
                    
                    Cells(2, numCheck + 2).Select
                    
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy
                    Cells(2, numCheck + 1).Select
                    Selection.PasteSpecial
                    
                    numCheck = numCheck + 1
                    
                Else
                    'need to insert a column and then merge the top
                    Columns(numCheck).Select
                    Selection.Insert shift:=xlRight
                    Columns(numCheck + 1).Copy
                    Columns(numCheck).PasteSpecial
                    Cells(1, numCheck).Clear

                    'merge first row... numcheck & numcheck + 1
                    Range(Cells(1, numCheck), Cells(1, numCheck + 1)).Select
                    Selection.Merge
                   
                End If
                
                On Error Resume Next 'Fix inexplicable problem of formula not copying to new columns
                    marRow = Application.WorksheetFunction.Match("*Margin", Range("F:F"), 0)
                On Error GoTo errhandler
                
                If marRow > 0 Then
                    Cells(marRow, numCheck + 1).Select
                    Selection.AutoFill Destination:=Range(Cells(marRow, numCheck), Cells(marRow, numCheck + 1))
                End If
                
            Else
            
                numCheck = k
            
            End If
            
            Cells(Application.WorksheetFunction.Match("Customer", Range("F:F"), 0), numCheck).Value = customeR
            Cells(Application.WorksheetFunction.Match("Type", Range("F:F"), 0), numCheck).Value = tyP
            Cells(Application.WorksheetFunction.Match("Serial #", Range("F:F"), 0), numCheck).Value = sN
            Cells(Application.WorksheetFunction.Match("Long Serial #", Range("F:F"), 0), numCheck).Value = snLong
            Cells(Application.WorksheetFunction.Match("Order #", Range("F:F"), 0), numCheck).Value = orderNum
            Cells(Application.WorksheetFunction.Match("Ship Month", Range("F:F"), 0), numCheck).Value = montH
            Cells(Application.WorksheetFunction.Match("Sell Price", Range("F:F"), 0), numCheck).Value = sellPrice
            Cells(Application.WorksheetFunction.Match("Current Material", Range("F:F"), 0), numCheck).Value = matlCost
            Cells(Application.WorksheetFunction.Match("Current Eng Cost", Range("F:F"), 0), numCheck).Value = engCost
            Cells(Application.WorksheetFunction.Match("Current Assy Cost", Range("F:F"), 0), numCheck).Value = mfgCost
            Cells(Application.WorksheetFunction.Match("Commission", Range("F:F"), 0), numCheck).Value = evoCost
            Cells(Application.WorksheetFunction.Match("Current*Test*Cost", Range("F:F"), 0), numCheck).Value = testCost
            
        End If
    
    Next j
    
    ThisWorkbook.Activate
    Sheet1.Activate
    Application.CutCopyMode = False
    Sheet1.Cells.Select
    Selection.Columns.AutoFit
    
Next i
WIP.Close savechanges:=False

'open eng & mfg wkbks
For Each wB In Application.Workbooks
    If wB.Name Like "*" & engFile & "*.xls*" Then
        wB.Activate
        Set engWB = ActiveWorkbook
    End If
Next wB

On Error GoTo referr
If engWB Is Nothing Then
    Set engWB = Workbooks.Open(fileName:=engPath & engFile & "*.xls*", UpdateLinks:=True)
End If
'''
For Each wB In Application.Workbooks
    If wB.Name Like "*" & mfgFile & "*.xls*" Then
        wB.Activate
        Set mfgWB = ActiveWorkbook
    End If
Next wB

On Error GoTo referr
If mfgWB Is Nothing Then
    Set mfgWB = Workbooks.Open(fileName:=mfgPath & mfgFile & "*.xls*", UpdateLinks:=True)
End If
On Error GoTo errhandler

ThisWorkbook.Activate
Sheet1.Activate

'delete jobs that are not in aux wkbks
Range("F2").Select
Selection.End(xlToRight).Select
numCols = ActiveCell.Column

Sheet1.Calculate
booContinue = DeleteTooNew(numCols)

If Not booContinue Then
    GoTo booErr
End If

On Error Resume Next
For i = 7 To numCols

    Range("D1").Formula = "=Substitute(Address(1, " & i & ", 4), 1, " & """" & """" & ")"
    Sheet1.Calculate
    colLet = Range("D1").Value
    
    If Range(colLet & 2).Value > 0 Then
        Range(colLet & Application.WorksheetFunction.Match("*Margin (%)", Range("F:F"), 0)).Formula = _
        "=iferror((" & colLet & "10-" & colLet & "31)/" & colLet & "10,0)"
        'sold margin values are hardcoded ^^^ because of laziness
    End If
    
    Cells(Application.WorksheetFunction.Match("*Margin", Range("F:F"), 0), i).HorizontalAlignment = xlCenter
    
Next i
Range("D1").ClearContents
On Error GoTo errhandler

Range("F2").Select
Selection.End(xlDown).Select
ActiveCell.Value = Date & " Margin"
Range(ActiveCell.Row & ":" & ActiveCell.Row).NumberFormat = "0%"

Call ReportAttrib(newDataPoint)

If newDataPoint Then
'update the scorecards
    Sheet1.Calculate
    booContinue = CardUpdate(engWB, mfgWB)
    
    If Not booContinue Then
        GoTo booErr
    End If
End If

Call ArrangeWorksheets
'puts the tabs in alpha order

booErr:

Application.Calculation = xlCalculationAutomatic
On Error Resume Next
engWB.Close savechanges:=False
mfgWB.Close savechanges:=False
On Error GoTo errhandler

ThisWorkbook.Activate
Sheet1.Activate
Columns("A:C").EntireColumn.Hidden = True
Columns(5).EntireColumn.Hidden = True
Range("G3").Select
ActiveWindow.FreezePanes = True
Range("D1").Select
Application.ScreenUpdating = True

If booContinue = True Then
    If newDataPoint Then
        MsgBox "Project list and scorecards updated successfully"
    Else
        MsgBox "Project list updated successfully"
    End If
End If

Exit Sub

referr: MsgBox "Error opening the reference workbooks. Check..." & vbCrLf & vbCrLf & _
            "1. Column A values for paths & filenames are correct" & vbCrLf & _
            "2. There are no other workbooks with those names already open"
        Exit Sub

errhandler: ThisWorkbook.Activate
            Sheet1.Activate
            Range("D1").ClearContents
            MsgBox "Unspecified error"
            Application.ScreenUpdating = True

End Sub

Function CardUpdate(engWB As Workbook, mfgWB As Workbook) As Boolean

Dim sellPrice As Single, cost As Single, i As Integer, numCols As Integer, booContinue As Boolean

On Error GoTo errhandler

CardUpdate = False

ThisWorkbook.Activate
Sheet1.Activate
Sheet1.Calculate
Range("F2").Select
Selection.End(xlToRight).Select
numCols = ActiveCell.Column

booContinue = PopulateScorecards(numCols)

If Not booContinue Then
    Exit Function
End If

ThisWorkbook.Activate
Sheet1.Activate
Range("D1").Select
'Range("F36").Value = Date

CardUpdate = True
Exit Function

errhandler: MsgBox "Error in CardUpdate function."


End Function

Function PopulateScorecards(numCols As Integer) As Boolean

Dim customerName As String, sN As String, orderDate As Date, shipDate As String, sellPrice As Single
Dim soldMarginDiff As Single, projMarginDiff As Single, engRel As Single, mfgRel As Single
Dim soldEngHrs As Single, soldMfgHrs As Single, soldMargin As Single, soldMatl As Single
Dim projMargin As Single, curMatl As Single, swRel As Single, sameDay As Boolean, testHrs As Single
Dim matlVar As Single, snLong As String, projLeader As String, j As Integer, lastDate As Date
Dim newSheet As Boolean, i As Integer, prodLine As String, wSheet As Worksheet, newChartRow As Integer
Dim indexBurt As Integer, indexCarr As Integer, indexMat As Integer, sheetName As String, dayMsg As Boolean
Dim curEngDesHrs As Single, curEngTestHrs As Single, remEngDesHrs As Single, remEngTestHrs As Single
Dim curMfgAssyHrs As Single, curMfgTestHrs As Single, remMfgAssyHrs As Single, remMfgTestHrs As Single
Dim wSheetNames() As String, x As Integer, wSheetName As Variant, secondMach As Boolean, xMach As Boolean
Dim invalidChars(6) As String, varVar As Variant

''''''''''''hardcoded'''''''''
invalidChars(0) = "\"
invalidChars(1) = "/"
invalidChars(2) = "*"
invalidChars(3) = "?"
invalidChars(4) = ":"
invalidChars(5) = "["
invalidChars(6) = "]"
''''''''''''''''''''''''''''''

On Error GoTo errhandler
Application.ScreenUpdating = False

PopulateScorecards = False

'find which columns are product line delineators
indexBurt = Application.WorksheetFunction.Match("Burt", Range("1:1"), 0)
indexCarr = Application.WorksheetFunction.Match("Carr", Range("1:1"), 0)
indexMateer = Application.WorksheetFunction.Match("Mateer", Range("1:1"), 0)
dayMsg = True 'don't remember what this is for???
For Each Sheet In ThisWorkbook.Sheets
    Sheet.Visible = xlSheetVisible
Next Sheet

For i = 7 To numCols
    
    On Error Resume Next
    
    customerName = CheckForLetError("Customer", i)
    If customerName = "" Then
        GoTo blankSN
    End If
    sN = CheckForLetError("Serial #", i)
    snLong = CheckForLetError("Long Serial #", i)
    projLeader = CheckForLetError("Project Leader", i)
    orderDate = 0
    shipDate = CheckForLetError("Ship Month", i)
    sellPrice = CheckForNumError("Sell Price", i)
    soldMarginDiff = CheckForNumError("Sold Margin ($)", i)
    projMarginDiff = sellPrice - CheckForNumError("Projected Cost", i)
    If projMarginDiff = sellPrice Then
        projMarginDiff = 0
    End If
    
    engRel = CheckForNumError("Engineering Released", i)
    If engRel > 1 Then
        engRel = engRel / 100
    End If
    
    mfgRel = CheckForNumError("Assembly Complete", i)
    If mfgRel > 1 Then
        mfgRel = mfgRel / 100
    End If
    
    swRel = CheckForNumError("Software Complete", i)
    If swRel > 1 Then
        swRel = swRel / 100
    End If

    soldEngHrs = CheckForNumError("Sold Eng Hrs", i)
    soldMfgHrs = CheckForNumError("Sold Assy Hrs", i)
    soldMargin = CheckForNumError("Sold Margin (%)", i)
    soldMatl = CheckForNumError("Sold Material", i)
    curEngDesHrs = CheckForNumError("Current Eng Hrs", i)
    'curEngTestHrs = CheckForNumError("Current Eng Test Hrs", i)
    remEngDesHrs = CheckForNumError("# Eng Hrs Remain", i)
    remEngTestHrs = CheckForNumError("# Eng Test Hrs Remain", i)
    curMfgAssyHrs = CheckForNumError("Current Assy Hrs", i)
    'curMfgTestHrs = CheckForNumError("Current Assy Test Hrs", i)
    remMfgAssyHrs = CheckForNumError("# Assy Hrs Remain", i)
    remMfgTestHrs = CheckForNumError("# Assy Test Hrs Remain", i)
    projMargin = CheckForMargError(i)
    curMatl = CheckForNumError("Current Material", i)
    matlVar = CheckForNumError("Material Variance", i)
    testHrs = CheckForNumError("Current Testing Hrs", i)

    On Error GoTo errhandler
    
    If i < indexCarr Then
        prodLine = "Burt"
    Else
        prodLine = "Carr"
    End If
    If i >= indexMateer Then
        prodLine = "Mateer"
        snLong = sN
        'mateer doesn't use the long S/N
    End If
    
    If Trim(UCase(customerName)) <> "STOCK" Then 'machine has been sold
        sheetName = prodLine & "-" & Left(Replace(customerName, " ", ""), WorksheetFunction.Min(5, Len(Replace(customerName, " ", ""))))
    Else 'machine hasn't been sold
        sheetName = prodLine & "-" & sN
    End If
    
    For Each varVar In invalidChars
        sheetName = Replace(sheetName, varVar, " ")
    Next
    
    x = 0 'create array of sheet names for comparison
    For Each wSheet In ThisWorkbook.Worksheets
        ReDim Preserve wSheetNames(x)
        wSheetNames(x) = wSheet.Name
        x = x + 1
    Next
    
    secondMach = False 'initial state
    xMach = False 'initial state
    For Each wSheetName In wSheetNames
        If wSheetName = sheetName Then 'already a machine tab with this name (unnumbered)
            secondMach = True
        ElseIf wSheetName Like sheetName & "(" & "#" & ")" Then 'already a machine tab with this name (numbered)
            xMach = True
        End If
    Next
    
    If secondMach Then 'one machine for this customer already
        ThisWorkbook.Worksheets(sheetName).Activate
        If Range("D2").Value <> snLong Then
            ActiveSheet.Name = sheetName & "(1)"
            sheetName = sheetName & "(2)"
        End If
    ElseIf xMach Then 'some number >1 of machines for this customer already
        x = 0
        newSheet = True
        For Each wSheet In ThisWorkbook.Worksheets
            If wSheet.Name Like sheetName & "(" & "#" & ")" Then
                wSheet.Activate
                If UCase(Range("D2").Value) = UCase(snLong) Then
                    newSheet = False
                    sheetName = wSheet.Name
                    Exit For
                Else
                    x = x + 1
                End If
            End If
        Next
        If newSheet Then
            sheetName = sheetName & "(" & x + 1 & ")"
        End If
    End If
    
    newSheet = True
    For Each wSheet In ThisWorkbook.Worksheets
        If wSheet.Name = sheetName Then
            newSheet = False
        End If
    Next wSheet
    
    If newSheet Then
        Sheet5.Copy After:=Sheet1
        ActiveSheet.Name = sheetName
    Else
        ThisWorkbook.Worksheets(sheetName).Activate
    End If

    'moved up here because we need to know the last
    'date the scorecard was updated
    Range("B" & Application.WorksheetFunction.Match("Date", Range("B:B"), 0)).Select
    Selection.End(xlDown).Select
    If newSheet Or ActiveCell.Row > 500 Then
        newChartRow = Application.WorksheetFunction.Match("Date", Range("B:B"), 0) + 1
        lastDate = 0
    Else
        Range("B" & Application.WorksheetFunction.Match("Date", Range("B:B"), 0)).Select
        Selection.End(xlDown).Select
        lastDate = ActiveCell.Value
        newChartRow = ActiveCell.Row + 1
    End If
    
    sameDay = False

    If Date - lastDate < 4 Then 'John requested 'same day' popup apply to 3 days
        sameDay = True
    End If

    ActiveSheet.Unprotect

    Range("D1").Value = customerName
    Range("D2").Value = snLong
    Range("D3").Value = projLeader
    If orderDate > 0 Then
        Range("D5").Value = orderDate
    Else
        Range("D5").Value = "NONE"
    End If

    Range("D6").Value = shipDate
    Range("D8").Value = sellPrice
    Range("D9").Value = soldMarginDiff
    Range("D10").Value = projMarginDiff
    Range("D12").Value = engRel
    Range("D13").Value = swRel
    Range("D14").Value = mfgRel
    Range("B17").Value = soldEngHrs
    Range("D17").Value = soldMfgHrs
    Range("F17").Value = soldMargin
    Range("G17").Value = soldMatl

    If newSheet Then
        For j = 2 To 7
            Cells(18, j).Value = "-"
        Next j
    ElseIf sameDay And dayMsg = True Then
        MsgBox "These cards were updated within the last 3 days. This update will overwrite the most recent one."
        dayMsg = False
    Else
        'change dashes to 0 if need be
        For j = 2 To 7
            If Cells(19, j).Value = "-" Then
                Cells(19, j).Value = 0
            End If
        Next j
        
        If Not sameDay Then
            Range("B18").Value = Range("B19").Value + Range("B20").Value
            Range("C18").Value = Range("C19").Value + Range("C20").Value
            Range("D18").Value = Range("D19").Value + Range("D20").Value
            Range("E18").Value = Range("E19").Value + Range("E20").Value
            Range("F18").Value = Range("F19").Value
            If Range("F18").Value = 0 Then
                Range("F18").Value = "-"
            End If
            Range("G18").Value = Range("G19").Value
            If Range("G18").Value = 0 Then
                Range("G18").Value = "-"
            End If
        End If
    End If

    Range("B19").Value = curEngDesHrs
    Range("B20").Value = remEngDesHrs
    'Range("C19").Value = curEngTestHrs
    Range("C20").Value = remEngTestHrs
    Range("D19").Value = curMfgAssyHrs
    Range("D20").Value = remMfgAssyHrs
    'Range("E19").Value = curMfgTestHrs
    Range("E20").Value = remMfgTestHrs
    Range("F19").Value = projMargin
    Range("G19").Value = Application.Max(soldMatl, curMatl) + matlVar
    Range("G21").Value = matlVar

    'testing hours
    If Left(Range("D2").Value, 1) = "C" Then 'carr
        Range("C19").Value = testHrs
        Range("E19").Value = "-"
    Else 'mateer/burt
        Range("E19").Value = testHrs
        Range("C19").Value = "-"
    End If
    
    If Not sameDay Then
        With Range("B" & newChartRow)
            .Value = Date
            .NumberFormat = "m/d/yy"
        End With
        
        With Range("C" & newChartRow)
            .Value = projMargin
            .NumberFormat = "0%"
        End With
        
        With Range("D" & newChartRow)
            .Value = soldMargin
            .NumberFormat = "0%"
        End With
    End If
    
    ActiveSheet.Cells.Select
    Selection.Columns.AutoFit
    Range("A1").Select
    ActiveSheet.Protect
    
blankSN:
    Sheet1.Activate
    
Next i

Sheet5.Visible = xlSheetHidden

PopulateScorecards = True

Exit Function

errhandler: MsgBox "Error in PopulateScorecards function."

End Function

Function ReportAttrib(newDataPoint)

Dim wSheet As Worksheet

For Each wSheet In ThisWorkbook.Worksheets
    wSheet.PageSetup.LeftFooter = "Program code written by: England, Tyler (PSA-CLW)"
    If newDataPoint Then
        If Application.UserName > "" Then
            wSheet.PageSetup.RightFooter = "Scorecard report generated by: " & Application.UserName
        Else
            wSheet.PageSetup.RightFooter = "Scorecard report generated by: UNKNOWN"
        End If
    End If
Next wSheet

End Function

Function CheckForMargError(i As Integer) As Single

Dim x As Boolean, rowNum As Integer

On Error GoTo errhandler

rowNum = Application.WorksheetFunction.Match("*Margin", Range("F:F"), 0)

x = IsError(Cells(rowNum, i))

If x = True Then
    CheckForMargError = 0
Else
    CheckForMargError = Cells(rowNum, i).Value

End If

Exit Function

errhandler: MsgBox "Error in CheckForNumError function."

End Function

Function CheckForNumError(searchTerm As String, i As Integer) As Single

Dim x As Boolean

On Error GoTo errhandler

x = IsError(Cells(Application.WorksheetFunction.Match(searchTerm, Range("F:F"), 0), i))

If x = True Then
    CheckForNumError = 0
End If

x = IsNumeric(Cells(Application.WorksheetFunction.Match(searchTerm, Range("F:F"), 0), i))

If x = False Then
    CheckForNumError = 0
Else
    CheckForNumError = Cells(Application.WorksheetFunction.Match(searchTerm, Range("F:F"), 0), i).Value
End If

Exit Function

errhandler: checkfornunerror = 0
            MsgBox "Error in CheckForNumError function."

End Function

Function CheckForLetError(searchTerm As String, i As Integer) As String

Dim x As Boolean

On Error GoTo errhandler

x = IsError(Cells(Application.WorksheetFunction.Match(searchTerm, Range("F:F"), 0), i))

If x = True Then
    CheckForLetError = "ERROR"
Else
    CheckForLetError = Cells(Application.WorksheetFunction.Match(searchTerm, Range("F:F"), 0), i)
End If

Exit Function

errhandler: MsgBox "Error in CheckForLetError function."

End Function


Sub ArrangeWorksheets()

'put tabs in alphabetical order

Dim sCount As Integer, i As Integer, j As Integer
 
Application.ScreenUpdating = False
 
sCount = Worksheets.Count

For i = 1 To (sCount - 1)
    For j = i + 1 To sCount
        If Worksheets(j).Name < Worksheets(i).Name Then
            Worksheets(j).Move Before:=Worksheets(i)
        End If
    Next j
Next i

Sheet1.Activate

End Sub

Function DeleteTooNew(numCols As Integer) As Boolean

Dim i As Integer, j As Integer, k As Integer, x As Boolean, firstRow As String, prodLine As String
Dim indexBurt As Integer, indexCarr As Integer, indexMateer As Integer, wkbkMissing As String

On Error GoTo errhandler

DeleteTooNew = False

ThisWorkbook.Activate
Sheet1.Activate

indexBurt = Application.WorksheetFunction.Match("Burt", Range("1:1"), 0)
indexCarr = Application.WorksheetFunction.Match("Carr", Range("1:1"), 0)
indexMateer = Application.WorksheetFunction.Match("Mateer", Range("1:1"), 0)

'using "sold hrs" to test whether row is in PM/MFG workbook
For i = 7 To numCols
    
    x = False
    'not to be deleted
    
    If Not (IsNumeric(Cells(14, i).Value)) And Cells(2, i).Value > 0 Then
        'missing from PM workbook
        x = True
        wkbkMissing = "engineering manager"
    ElseIf Not (IsNumeric(Cells(21, i).Value)) And Cells(2, i).Value > 0 Then
        'missing from MFG workbook
        x = True
        wkbkMissing = "assembly/mfg manager"
    End If
    
    If x Then
        prodLine = Cells(1, i).Value
            
        If i = indexBurt Or i = indexCarr Or i = indexMateer Then
            'deleting the column with the product line name
            firstRow = Cells(1, i).Value
            j = Application.WorksheetFunction.Match(prodLine, Range("A:A"), 0)
        Else
            firstRow = "NONE"
        End If
        
        k = i
        Do While prodLine = ""
           prodLine = Cells(1, k).Value
           k = k - 1
        Loop
        
        MsgBox "The " & Application.WorksheetFunction.Proper(Cells(2, i).Value) & " machine (" & prodLine & _
            " S/N " & Cells(5, i).Value & ") is not being included " & _
            "because it doesn't appear in " & wkbkMissing & " workbook."
    
        Columns(i).Delete
    
        If firstRow <> "NONE" Then
            'put product line name back
            Cells(1, i).Formula = "=$A$" & j
        End If
        
        If i < indexCarr Then
            indexCarr = indexCarr - 1
        End If
        
        If i < indexMateer Then
            indexMateer = indexMateer - 1
        End If
        
        i = i - 1
        
    End If
Next i

DeleteTooNew = True
Exit Function

errhandler: MsgBox "Error in DeleteTooNew function."

End Function


Sub PrintScorecards()

Dim wSheet As Worksheet

Call ReportAttrib(False)

For Each wSheet In ThisWorkbook.Worksheets
    
    If wSheet.Name <> "A_New_Scorecard" And wSheet.Name <> "All Projects" And wSheet.Visible = xlSheetVisible Then
        
        wSheet.PrintOut from:=1, To:=1
        
    End If

Next wSheet

Sheet1.Activate

End Sub
