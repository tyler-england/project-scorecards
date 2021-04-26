Attribute VB_Name = "FuncsOtherWBs"
Option Explicit

Function GetFromENG(ByRef sCOs() As String) As Object
'''Returns ENG/MFG estimates/values for active projects
'''Error returns Nothing
    Set GetFromENG = Nothing
    On Error GoTo errhandler
    
    Dim oDictOut As Object, oDictCO As Object
    Dim sCO As String, sOldCO As String 'used if machine was STOCK and now has a CO
    Dim i As Integer, j As Integer, iRow As Integer, iEmpty As Integer
    Dim wbENG As Workbook, sWbName As String
    Dim vRng As Variant
    Dim bFound As Boolean, bWasOpen As Boolean

    Call GlobalVars
    Set oDictOut = CreateObject("Scripting.Dictionary")
    
    If InStr(wbDataENG, "/") > 0 Then
        sWbName = Right(wbDataENG, Len(wbDataENG) - InStrRev(wbDataENG, "/"))
    ElseIf InStr(wbDataENG, "\") > 0 Then
        sWbName = Right(wbDataENG, Len(wbDataENG) - InStrRev(wbDataENG, "\"))
    End If
    For Each vRng In Workbooks
        If vRng.Name = sWbName Then
            Set wbENG = vRng
            bWasOpen = True
            Exit For
        End If
    Next
    If wbENG Is Nothing Then Set wbENG = Workbooks.Open(wbDataENG)
    
    vRng = wbENG.Worksheets(1).Range("A1:BM500")
    If Not bWasOpen Then wbENG.Close savechanges:=False
    
    iRow = 2 'ignore header row
    Do While iEmpty < 3
        If vRng(iRow, 1) = 0 Then
            iEmpty = iEmpty + 1
        Else
            iEmpty = 0
            bFound = False 'indicates when found in existing card COs
            For j = 0 To UBound(sCOs)
                If vRng(iRow, 3) = sCOs(j) Then
                    bFound = True 'has existing card, CO field is correct
                    Exit For
                End If
            Next
            If Not bFound Then 'see if existing but has new CO (was using SN until now)
                For j = 0 To UBound(sCOs)
                    If UCase(vRng(iRow, 5)) Like "*" & UCase(sCOs(j)) & "*" Then 'has new CO
                        bFound = True
                        sOldCO = sCOs(j)
                    End If
                Next
            End If
            Set oDictCO = CreateObject("Scripting.Dictionary") 'store proj data in sCO dictionary
            oDictCO.Add "PL", vRng(iRow, 1)
            oDictCO.Add "Cust", vRng(iRow, 2)
            oDictCO.Add "CO", vRng(iRow, 3)
            oDictCO.Add "MO", vRng(iRow, 4)
            oDictCO.Add "SN", vRng(iRow, 5)
            oDictCO.Add "SoldHrsME", vRng(iRow, 6)
            oDictCO.Add "RemHrsME", vRng(iRow, 9)
            oDictCO.Add "SoldHrsEE", vRng(iRow, 10)
            oDictCO.Add "RemHrsEE", vRng(iRow, 13)
            oDictCO.Add "SoldHrsSW", vRng(iRow, 14)
            oDictCO.Add "RemHrsSW", vRng(iRow, 17)
            oDictCO.Add "SoldHrsET", vRng(iRow, 18)
            oDictCO.Add "RemHrsET", vRng(iRow, 21)
            oDictCO.Add "MatlVar", vRng(iRow, 22)
            oDictCO.Add "%ME", vRng(iRow, 23)
            oDictCO.Add "%EE", vRng(iRow, 24)
            oDictCO.Add "%SW", vRng(iRow, 25)
            oDictCO.Add "LeadME", vRng(iRow, 26)
            oDictCO.Add "LeadEE", vRng(iRow, 27)
            oDictCO.Add "CmtENG", vRng(iRow, 28)
            oDictCO.Add "SoldHrsMA", vRng(iRow, 29)
            oDictCO.Add "RemHrsMA", vRng(iRow, 32)
            oDictCO.Add "SoldHrsEA", vRng(iRow, 33)
            oDictCO.Add "RemHrsEA", vRng(iRow, 36)
            oDictCO.Add "SoldHrsAT", vRng(iRow, 37)
            oDictCO.Add "RemHrsAT", vRng(iRow, 40)
            oDictCO.Add "%MFG", vRng(iRow, 41)
            oDictCO.Add "%TS", vRng(iRow, 42)
            oDictCO.Add "LeadMA", vRng(iRow, 45)
            oDictCO.Add "LeadEA", vRng(iRow, 46)
            oDictCO.Add "CmtMFG", vRng(iRow, 47)
            oDictCO.Add "SellPrice", vRng(iRow, 48)
            oDictCO.Add "SalesRep", vRng(iRow, 49)
            oDictCO.Add "DateSell", vRng(iRow, 50)
            oDictCO.Add "DateShip", vRng(iRow, 51)
            oDictCO.Add "Destination", vRng(iRow, 52)
            oDictCO.Add "DateFAT", vRng(iRow, 53)
            oDictCO.Add "PM", vRng(iRow, 54)
            oDictCO.Add "SoldMat", vRng(iRow, 55)
            oDictCO.Add "SoldMar", vRng(iRow, 56)
            
            If sOldCO <> "" Then 'use that so proper scorecard gets updated --> next time the CO will be right
                oDictOut.Add sOldCO, oDictCO
                sOldCO = ""
            Else
                oDictOut.Add CStr(vRng(iRow, 3)), oDictCO 'add sCO dictionary to overall dictionary of CO info
            End If
            
            Set oDictCO = Nothing
        End If
        iRow = iRow + 1
    Loop
    
    Set GetFromENG = oDictOut
    Exit Function
errhandler:
    MsgBox "Errrrr"
End Function

Function GetActuals(ByRef sCOs() As String) As Object
'''Returns Actual hours/mat'l/Lead Times for projects
'''Error returns Nothing
    Set GetActuals = Nothing
    'On Error GoTo errhandler
    
    Dim oDictOut As Object, oDictCO As Object
    Dim i As Integer, j As Integer, iEmpty As Integer
    Dim sPartsAll() As String, sParts20(19) As String
    Dim wbCO As Workbook, sWbName As String
    Dim vRng As Variant, vParts As Variant
    Dim bWasOpen As Boolean
    
    Call GlobalVars
    Set oDictOut = CreateObject("Scripting.Dictionary")

    If InStr(wbDataCO, "/") > 0 Then
        sWbName = Right(wbDataCO, Len(wbDataCO) - InStrRev(wbDataCO, "/"))
    ElseIf InStr(wbDataCO, "\") > 0 Then
        sWbName = Right(wbDataCO, Len(wbDataCO) - InStrRev(wbDataCO, "\"))
    End If
    For Each vRng In Workbooks
        If vRng.Name = sWbName Then
            Set wbCO = vRng
            bWasOpen = True
            Exit For
        End If
    Next
    If wbCO Is Nothing Then Set wbCO = Workbooks.Open(wbDataCO)
        
    With wbCO.Worksheets("CO List")
        .Range("A2:A500").ClearContents
        .Range("A2").Resize(UBound(sCOs) + 1) = Application.Transpose(sCOs)
    End With
    
    Application.Run "'" & wbCO.Name & "'!RefreshData" 'update co data (run macro)
    vRng = wbCO.Worksheets("Summary").Range("A1:J500").Value2
    If Not bWasOpen Then wbCO.Close savechanges:=False
    
    i = 3
    Do While iEmpty < 3 'go through summary --> for each, build dict and add to master dict
        If vRng(i, 1) = 0 Then
            iEmpty = iEmpty + 1
        Else
            iEmpty = 0
            Set oDictCO = CreateObject("Scripting.Dictionary")
            oDictCO.Add "Matl", Int(vRng(i, 2))
            oDictCO.Add "HrsME", Int(vRng(i, 3))
            oDictCO.Add "HrsEE", Int(vRng(i, 4))
            oDictCO.Add "HrsSW", Int(vRng(i, 5))
            oDictCO.Add "HrsET", 0 ' we don't track ENG test hrs yet (??)
            oDictCO.Add "HrsMA", Int(vRng(i, 6))
            oDictCO.Add "HrsEA", Int(vRng(i, 7))
            oDictCO.Add "HrsTS", Int(vRng(i, 8))
            
            vParts = Split(vRng(i, 10), ";;")
            For j = 0 To UBound(vParts) 'get only the latest 20
                If j > 19 Then Exit For
                sParts20(j) = vParts(j)
            Next
            oDictCO.Add "Parts", sParts20
            
            oDictOut.Add CStr(vRng(i, 1)), oDictCO 'add CO dictionary to dictionary of all COs
            
            Set oDictCO = Nothing
        End If
        i = i + 1
    Loop
    
    Set GetActuals = oDictOut
    
    Exit Function
errhandler:
    MsgBox "Err actuals"
End Function

Function GetDocs(ByRef sCOs() As String) As Object
'''Returns status of all documents related to projects
'''Error returns Nothing
    Set GetDocs = Nothing
    On Error GoTo errhandler
    
    Dim oDictOut As Object, oDictCO As Object
    Dim wbDocs As Workbook, sWbName As String
    Dim i As Integer, j As Integer, iEmpty As Integer
    Dim bCol As Boolean, bWasOpen As Boolean
    Dim vRng As Variant
    
    Call GlobalVars
    Set oDictOut = CreateObject("Scripting.Dictionary")
    
    If InStr(wbDocTracker, "/") > 0 Then
        sWbName = Right(wbDocTracker, Len(wbDocTracker) - InStrRev(wbDocTracker, "/"))
    ElseIf InStr(wbDocTracker, "\") > 0 Then
        sWbName = Right(wbDocTracker, Len(wbDocTracker) - InStrRev(wbDocTracker, "\"))
    End If
    For Each vRng In Workbooks
        If vRng.Name = sWbName Then
            Set wbDocs = vRng
            bWasOpen = True
            Exit For
        End If
    Next
    If wbDocs Is Nothing Then Set wbDocs = Workbooks.Open(wbDocTracker)
    
    vRng = wbDocs.Worksheets(1).Range("C1:CZ100").Value2
    If Not bWasOpen Then wbDocs.Close savechanges:=False
    For i = 0 To UBound(sCOs)
        j = 1
        bCol = False
        iEmpty = 0
        Do While iEmpty < 3
            If vRng(3, j) = 0 Then
                iEmpty = iEmpty + 1
            Else
                iEmpty = 0
                If vRng(3, j) = sCOs(i) Then
                    bCol = True
                    Exit Do
                ElseIf vRng(2, j) Like "*" & sCOs(i) & "*" Then
                    bCol = True
                    Exit Do
                End If
            End If
            j = j + 1
        Loop
        
        If bCol Then 'column found for current job
            Set oDictCO = CreateObject("Scripting.Dictionary") 'Part# // Vault stat // ENG/Date/comment
            oDictCO.Add "Build", vRng(10, j) & "//" & vRng(11, j)
            oDictCO.Add "Test", vRng(12, j) & "//" & vRng(13, j)
            oDictCO.Add "OpMan", vRng(14, j) & "//" & vRng(15, j)
            oDictCO.Add "TransOpMan", vRng(16, j) & "//" & vRng(17, j)
            oDictCO.Add "cGMP", vRng(18, j) & "//" & vRng(19, j)
            oDictCO.Add "Comp", vRng(20, j) & "//" & vRng(21, j)
            oDictCO.Add "SAT", vRng(22, j) & "//" & vRng(23, j)
            oDictCO.Add "IQOQ", vRng(24, j) & "//" & vRng(25, j)
            oDictCO.Add "DMR", vRng(26, j)
            oDictCO.Add "URS", vRng(28, j)
            oDictCO.Add "Process", vRng(29, j) & "//" & vRng(30, j)
            oDictCO.Add "PIDTags", vRng(31, j) & "//" & vRng(32, j)
            oDictCO.Add "GA", vRng(33, j) & "//" & vRng(34, j) & "//" & vRng(35, j)
            oDictCO.Add "MechA", vRng(36, j) & "//" & vRng(37, j) & "//" & vRng(38, j)
            oDictCO.Add "PID", vRng(39, j) & "//" & vRng(40, j) & "//" & vRng(41, j)
            oDictCO.Add "ContPanelDWG", vRng(42, j) & "//" & vRng(43, j) & "//" & vRng(44, j)
            oDictCO.Add "ContPanelSchem", vRng(45, j) & "//" & vRng(46, j) & "//" & vRng(47, j)
            oDictCO.Add "DrivePanelDWG", vRng(48, j) & "//" & vRng(49, j) & "//" & vRng(50, j)
            oDictCO.Add "DrivePanelSchem", vRng(51, j) & "//" & vRng(52, j) & "//" & vRng(53, j)
            oDictCO.Add "PneumInt", vRng(54, j) & "//" & vRng(55, j) & "//" & vRng(56, j)
            oDictCO.Add "PneumSchem", vRng(57, j) & "//" & vRng(58, j) & "//" & vRng(59, j)
            oDictCO.Add "ElecInt", vRng(60, j) & "//" & vRng(61, j) & "//" & vRng(62, j)
            oDictCO.Add "FRS", vRng(63, j) & "//" & vRng(64, j) & "//" & vRng(65, j)
            oDictCO.Add "DDS", vRng(66, j) & "//" & vRng(67, j) & "//" & vRng(68, j)
            oDictCO.Add "Alarms", vRng(69, j) & "//" & vRng(70, j) & "//" & vRng(71, j)
            oDictCO.Add "IO", vRng(72, j) & "//" & vRng(73, j) & "//" & vRng(74, j)
            oDictCO.Add "Interlocks", vRng(75, j) & "//" & vRng(76, j) & "//" & vRng(77, j)
            oDictCO.Add "CIP", vRng(78, j) & "//" & vRng(79, j) & "//" & vRng(80, j)
            oDictCO.Add "SIP", vRng(81, j) & "//" & vRng(82, j) & "//" & vRng(83, j)
            oDictCO.Add "VFD", vRng(84, j) & "//" & vRng(85, j) & "//" & vRng(86, j)
            oDictCO.Add "Screens", vRng(87, j) & "//" & vRng(88, j) & "//" & vRng(89, j)
            oDictCO.Add "Software", vRng(90, j) & "//" & vRng(91, j) & "//" & vRng(92, j)
            
            oDictOut.Add sCOs(i), oDictCO
            
            Set oDictCO = Nothing
        End If
    Next
    
    Set GetDocs = oDictOut
    Exit Function
errhandler:
    MsgBox "Docs prob"
End Function

Function UpdateEngMfgApps(oDictActs As Object) As String()
'''Returns changed COs if workbook has updated (and there are changes)
    
    Dim sOut() As String, sCO As String, sKeys(6) As String
    Dim wbENG As Workbook, bWasOpen As Boolean
    Dim i As Integer, j As Integer, x As Integer, iCols(6) As Integer
    Dim var As Variant
    
    On Error GoTo errhandler
    
    iCols(0) = 9
    iCols(1) = 13
    iCols(2) = 17
    iCols(3) = 21
    iCols(4) = 32
    iCols(5) = 36
    iCols(6) = 40
    
    sKeys(0) = "HrsME"
    sKeys(1) = "HrsEE"
    sKeys(2) = "HrsSW"
    sKeys(3) = "HrsET"
    sKeys(4) = "HrsMA"
    sKeys(5) = "HrsEA"
    sKeys(6) = "HrsTS"
    
    If InStr(wbDataENG, "/") > 0 Then
        sWbName = Right(wbDataENG, Len(wbDataENG) - InStrRev(wbDataENG, "/"))
    ElseIf InStr(wbDataENG, "\") > 0 Then
        sWbName = Right(wbDataENG, Len(wbDataENG) - InStrRev(wbDataENG, "\"))
    End If
    For Each var In Workbooks
        If var.Name = sWbName Then
            Set wbENG = var
            bWasOpen = True
            Exit For
        End If
    Next
    If wbENG Is Nothing Then Set wbENG = Workbooks.Open(wbDataENG)
    
    i = 4 'first row in ENG wb
    ReDim sOut(0)
    With wbENG.Worksheets(1)
        Do While .Range("C" & i).Value > 0
            sCO = .Range("C" & i).Value
            If oDictActs.Exists(sCO) Then
                For j = 0 To UBound(iCols)
                    If .Cells(i, j - 1).Value <> oDictActs(sCO).Item(sKeys(j)) Then
                        .Cells(i, j).Value = .Cells(i, j - 2).Value - oDictActs(sCO).Item(sKeys(j))
                        .Cells(i, j - 1).Value = oDictActs(sCO).Item(sKeys(j))
                        ReDim Preserve sOut(x)
                        sOut(x) = "• " & sCO
                        x = x + 1
                    End If
                Next
            End If
            wbENG.Save
            i = i + 1
        Loop
    End With
    
    If Not bWasOpen Then wbENG.Close savechanges:=False
    
    UpdateEngMfgApps = sOut
    Exit Function
errhandler:
    MsgBox "Error updating ENG/MFG/APPS workbook"
End Function
'
'Function GetForENG() As Object
''''Returns dictionary for updating ENG/MFG workbook w/ current project values
''''Error returns Nothing
'    Dim oDictOut As Object, oDictCO As Object
'    Dim sCO As String
'    Dim i As Integer, j As Integer, iDiff As Integer
'    Dim wSheet As Worksheet
'    Set oDictOut = CreateObject("Scripting.Dictionary")
'    For i = 0 To UBound(sCOs)
'        Set wSheet = Nothing
'        sCO = sCOs(i)
'        If i > 0 Then
'            If sCO = GetCO(ThisWorkbook.Worksheets(i + iDiff)) Then
'                Set wSheet = ThisWorkbook.Worksheets(i = iDiff)
'            End If
'        End If
'        If wSheet Is Nothing Then 'i=0 or idiff didn't work
'            For j = 1 To ThisWorkbook.Worksheets.Count
'                If sCOs(i) = GetCO(ThisWorkbook.Worksheets(j)) Then
'                    Set wSheet = ThisWorkbook.Worksheets(j)
'                    iDiff = j - i
'                    Exit For
'                End If
'            Next
'        End If
'        If wSheet Is Nothing Then 'should really never happen
'            MsgBox "Unable to find the scorecard worksheet for " & sCO(i) & ". No cards were updated.", vbExclamation
'            Exit Function
'        End If
'        Set oDictCO = CreateObject("Scripting.Dictionary")
'        With wSheet
'            oDictCO.Add "HrsME", .Range("J5").Value
'            oDictCO.Add "HrsEE", .Range("L5").Value
'            oDictCO.Add "HrsSW", .Range("N5").Value
'
'        End With
'        oDictOut.Add sCO(i), oDictCO
'        Set oDictCO = Nothing
'    Next
'    Set GetForENG = oDictOut
'    Exit Function
'errhandler:
'    MsgBox "Error in GetForENG function"
'End Function
