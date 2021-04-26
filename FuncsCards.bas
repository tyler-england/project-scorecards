Attribute VB_Name = "FuncsCards"
Option Explicit

Function CreateNewCards(ByRef sNewCOs() As String) As Boolean
'''Makes new cards & names them appropriately
'''Error returns False
    Dim sCO As String, sSheetName As String
    Dim i As Integer, j As Integer, k As Integer
    Dim bFound As Boolean
    Dim wsBlank As Worksheet, vWS As Variant
    
    On Error GoTo errhandler
    
    Set wsBlank = ThisWorkbook.Worksheets("#BLANK")
    wsBlank.Visible = True
    
    For i = 0 To UBound(sNewCOs)
        sCO = sNewCOs(i)
        sSheetName = sCO
        j = 1
        For Each vWS In ThisWorkbook.Worksheets
            If UCase(vWS.Name) = sSheetName Then
                bFound = True
                Exit For
            End If
        Next
        
        Do While bFound
            bFound = False
            sSheetName = sCO & "-" & CStr(j)
            j = i + 1
            For Each vWS In ThisWorkbook.Worksheets
                If UCase(vWS.Name) = sSheetName Then
                    bFound = True
                    Exit For
                End If
            Next
            If j > 100 Then 'something wrong
                sSheetName = "???" & Format(Now, "HH,MM,SS") & CStr(Rnd)
            End If
        Loop
        
        k = ThisWorkbook.Worksheets.Count
        wsBlank.Copy after:=ThisWorkbook.Worksheets(k)
        With ThisWorkbook.Worksheets(k + 1)
            .Name = sSheetName
            If sCO Like "######" Then 'actual CO
                .Range("C4").Value = sCO
            Else 'prob is SN
                .Range("C3").Value = sCO
            End If
        End With
    Next
    
    wsBlank.Visible = False
    CreateNewCards = True
errhandler:
End Function

Function UpdateCards(ByRef sCOs() As String, ByRef oDictENG As Object, ByRef oDictActs As Object, ByRef oDictDocs As Object) As Boolean
'''Updates cards for given COs
'''Error returns False
    Dim sCO As String, sSkipped() As String
    Dim bENG As Boolean, bACT As Boolean, bDOC As Boolean
    Dim i As Integer, x As Integer, iCount As Integer
    Dim wsCard As Worksheet
    Dim bContinue As Boolean
    
    'on error goto errhandler
    ReDim sSkipped(0)
    iCount = ThisWorkbook.Worksheets.Count
    For i = 4 To iCount
        Application.ScreenUpdating = True
        Application.StatusBar = "Updating scorecard " & CStr(i - 3) & "/" & CStr(iCount - 3) & "..."
        Application.ScreenUpdating = False
        Set wsCard = ThisWorkbook.Worksheets(i)
        sCO = GetCO(wsCard)
        bContinue = UpdateFormulas(wsCard)

        bENG = False
        bACT = False
        bDOC = False
        If Not bContinue Then Exit Function

        If oDictENG.Exists(sCO) Then
            bContinue = UpdateENG(oDictENG(sCO), wsCard)
            bENG = True
            If Not bContinue Then Exit Function
        End If
        If oDictActs.Exists(sCO) Then
            bContinue = UpdateActs(oDictActs(sCO), wsCard)
            bACT = True
            If Not bContinue Then Exit Function
        End If
        If oDictDocs.Exists(sCO) Then
            bContinue = UpdateDocs(oDictDocs(sCO), wsCard)
            bDOC = True
            If Not bContinue Then Exit Function
        End If
        'update project trend & margin
        If Not bENG Or Not bACT Or Not bDOC Then 'something failed / didn't exist
            ReDim Preserve sSkipped(x)
            sSkipped(x) = "· " & ThisWorkbook.Worksheets(i).Name & " - "
            If Not bENG Then sSkipped(x) = sSkipped(x) & "ENG data, "
            If Not bACT Then sSkipped(x) = sSkipped(x) & "XA data, "
            If Not bDOC Then sSkipped(x) = sSkipped(x) & "Docs data, "
            sSkipped(x) = Left(sSkipped(x), Len(sSkipped(x)) - 2)
            x = x + 1
        End If
    Next
    
    Call RearrangeCards(oDictENG) 'rename & alphabetize
    
    If x > 0 Then
        MsgBox "Some card updates were incomplete. The following cards are missing data:" & _
                vbCrLf & vbCrLf & Join(sSkipped, vbCrLf)
    End If
    
    UpdateCards = True
    Exit Function
errhandler:
    MsgBox "Error updating the scorecards" & vbCrLf & vbCrLf & "Err: " & Err.Number & vbCrLf & Err.Description
End Function

Function UpdateFormulas(wsCard As Worksheet) As Boolean
'''Updates "last week" values, graph values, etc. before they are updated for new week
    On Error GoTo errhandler
    
    Dim i As Integer, iRowFrom As Integer, iRowTo As Integer
    Dim dLatest As Date, dNew As Date, sVal As String
    
    iRowFrom = 10 'used + to use
    iRowTo = 13 'row that shows "last week"
    
    For i = 6 To 23 'move the values
        wsCard.Cells(iRowTo, i).Value = wsCard.Cells(iRowFrom, i).Value
    Next
    
    dNew = Date
    If Weekday(Date) > 4 Then
        dNew = Date - Weekday(Date) - 4
    ElseIf Weekday(Date) < 4 Then
        dNew = Date - 7 - (4 - Weekday(Date))
    End If
    
    On Error Resume Next
        dLatest = WorksheetFunction.Max(wsCard.Range("AI:AI"))
    On Error GoTo errhandler
    If dLatest >= dNew Then dNew = 0 'don't add any date (graph updated)
    
    If dNew <> 0 Then 'add to graph
        iRowTo = wsCard.Range("AI250").End(xlUp).Row + 1
        wsCard.Range("AI" & iRowTo).Value = dNew
        sVal = wsCard.Range("F" & iRowFrom).Value
        wsCard.Range("AJ" & iRowTo).Value = Mid(sVal, InStr(sVal, "(") + 1, InStr(sVal, ")") - InStr(sVal, "(") - 1)
        sVal = wsCard.Range("F" & iRowFrom - 1).Value
        wsCard.Range("AJ" & iRowTo).Value = Mid(sVal, InStr(sVal, "(") + 1, InStr(sVal, ")") - InStr(sVal, "(") - 1)
    End If
    
    UpdateFormulas = True
    
    Exit Function
errhandler:
    MsgBox "Error in UpdateFormulas function" & vbCrLf & vbCrLf & "Err: " & Err.Number & vbCrLf & Err.Description
End Function

Function UpdateENG(ByRef oDict As Object, ByRef wsCard As Worksheet) As Boolean
'''Fills in ENG/MFG/APP info on worksheet
    'On Error GoTo errhandler
    
    Dim siSell As Single, siCosts As Single, siMargin As Single
    Dim siRateENG As Single, siRateASSY As Single
    Dim sMargin As String
    
    With wsCard
        .Range("C2").Value = oDict("Cust")
        .Range("C3").Value = oDict("SN")
        .Range("C4").Value = oDict("CO")
        .Range("C5").Value = oDict("MO")
        .Range("C7").Value = oDict("SellPrice")
        .Range("C8").Value = oDict("Destination")
        .Range("C9").Value = oDict("DateSell")
        .Range("C10").Value = oDict("DateFAT")
        .Range("C11").Value = oDict("DateShip")
        .Range("C13").Value = oDict("PM")
        .Range("C14").Value = oDict("LeadME")
        .Range("C15").Value = oDict("LeadEE")
        .Range("C16").Value = oDict("LeadEA")
        .Range("C17").Value = oDict("LeadMA")
        .Range("C18").Value = oDict("SalesRep")
        
        .Range("F2").Value = Format(Now, "DD-MMM-YYYY")
        .Range("F3").Value = oDict("%ME")
        .Range("F4").Value = (oDict("%EE") + oDict("%SW")) / 2 'average --- do something else?
        .Range("F5").Value = oDict("%MFG")
        
        .Range("E16").Value = oDict("MatlVar")
        
        If InStr(UCase(.Range("Y14").Value), oDict("CmtENG")) = 0 Then .Range("Y14").Value = .Range("Y14").Value & "  " & oDict("CmtENG")
        If InStr(UCase(.Range("AB14").Value), oDict("CmtMFG")) = 0 Then .Range("AB14").Value = .Range("AB14").Value & "  " & oDict("CmtMFG")
'        If oDict("CmtENG") <> "" Then .Range("F16").Value = oDict("CmtENG") & "   "
'        If oDict("CmtMFG") <> "" Then .Range("F16").Value = .Range("F16").Value & oDict("CmtMFG")

        .Range("K5").Value = oDict("RemHrsME")
        .Range("M5").Value = oDict("RemHrsEE")
        .Range("O5").Value = oDict("RemHrsSW")
        .Range("Q5").Value = oDict("RemHrsET")
        .Range("S5").Value = oDict("RemHrsMA")
        .Range("U5").Value = oDict("RemHrsEE")
        .Range("W5").Value = oDict("RemHrsAT")

        siSell = .Range("C7").Value
        If UCase(oDict("PL")) = "BURT" Or UCase(oDict("PL")) = "MATEER" Then
            siRateENG = sRateBurtENG
            siRateASSY = sRateBurtASSY
            .Range("A1").Value = oDict("PL")
        ElseIf UCase(oDict("PL")) = "CARR" Then
            siRateENG = sRateCarrENG
            siRateASSY = sRateCarrASSY
            .Range("A1").Value = oDict("PL")
        Else
            siRateENG = WorksheetFunction.Max(sRateBurtENG, sRateCarrENG)
            siRateASSY = WorksheetFunction.Max(sRateBurtASSY, sRateCarrASSY)
        End If

        siCosts = .Range("G9").Value + siRateENG * WorksheetFunction.Sum(.Range("J9:Q9")) + siRateASSY * WorksheetFunction.Sum(.Range("R9:W9"))
        siMargin = siSell - siCosts
        If siSell <> 0 Then sMargin = Format(siMargin, "$0,000") & " (" & Int(siMargin * 100 / siSell) & "%)"
        .Range("F9").Value = sMargin
        .Range("H9").Value = oDict("SoldMat")
        .Range("J9").Value = oDict("SoldHrsME")
        .Range("L9").Value = oDict("SoldHrsEE")
        .Range("N9").Value = oDict("SoldHrsSW")
        .Range("P9").Value = oDict("SoldHrsET")
        .Range("R9").Value = oDict("SoldHrsMA")
        .Range("T9").Value = oDict("SoldHrsEA")
        .Range("V9").Value = oDict("SoldHrsAT")
    End With
    
    UpdateENG = True
    Exit Function
errhandler:
    MsgBox "Error updating ENG/MFG/APP info for " & wsCard.Name & vbCrLf & vbCrLf & "Program terminated"
End Function

Function UpdateActs(ByRef oDict As Object, ByRef wsCard As Worksheet) As Boolean
'''Fills in Actual Mat'l/Hrs info on worksheet
    On Error GoTo errhandler
    
    Dim iRow As Integer, i As Integer
    Dim vParts As Variant, vComps As Variant
    Dim sPN As String, sDesc As String, sDate As String, sPL As String, sMargin As String
    Dim siSell As Single, siCosts As Single, siMargin As Single
    Dim siRateENG As Single, siRateASSY As Single
    
    With wsCard
        .Range("J5").Value = oDict("HrsME")
        .Range("L5").Value = oDict("HrsEE")
        .Range("N5").Value = oDict("HrsSW")
        .Range("P5").Value = "0" '??
        .Range("R5").Value = oDict("HrsMA")
        .Range("T5").Value = oDict("HrsEA")
        .Range("V5").Value = oDict("HrsTS")
        
        siSell = .Range("C7").Value
        sPL = UCase(.Range("A1").Value)
        If sPL = "BURT" Or sPL = "MATEER" Then
            siRateENG = sRateBurtENG
            siRateASSY = sRateBurtASSY
        ElseIf sPL = "CARR" Then
            siRateENG = sRateCarrENG
            siRateASSY = sRateCarrASSY
        Else
            siRateENG = WorksheetFunction.Max(sRateBurtENG, sRateCarrENG)
            siRateASSY = WorksheetFunction.Max(sRateBurtASSY, sRateCarrASSY)
        End If
        siCosts = .Range("G10").Value + siRateENG * WorksheetFunction.Sum(.Range("J10:Q10")) + siRateASSY * WorksheetFunction.Sum(.Range("R10:W10"))
        siMargin = siSell - siCosts
        If siSell <> 0 Then sMargin = Format(siMargin, "$0,000") & " (" & Int(siMargin * 100 / siSell) & "%)"
        
        .Range("F10").Value = sMargin
        .Range("G10").Value = oDict("Matl")

        iRow = 24
        vParts = oDict("Parts")
        For i = 0 To UBound(vParts)
            If vParts(i) = "" Then Exit For
            sPN = vParts(i)
            sDesc = ""
            sDate = ""
            On Error Resume Next
            vComps = Split(vParts(i), "//")
            sPN = vComps(0)
            sDesc = vComps(1)
            sDate = vComps(2)
            On Error GoTo errhandler
            .Range("B" & iRow).Value = sPN
            .Range("C" & iRow).Value = sDesc
            .Range("F" & iRow).Value = sDate
            iRow = iRow + 1
        Next
    End With
    
    UpdateActs = True
    Exit Function
errhandler:
    MsgBox "Error updating Actual/XA info for " & wsCard.Name & vbCrLf & vbCrLf & "Program terminated"
End Function

Function UpdateDocs(ByRef oDict As Object, ByRef wsCard As Worksheet) As Boolean
'''Fills in Docs info on worksheet
    On Error GoTo errhandler
    
    Dim i As Integer, sDict As String, vComps As Variant
    With wsCard
        For i = 24 To 45
            sDict = oDict(.Range("G" & i).Value)
            If sDict <> "" Then 'doc has value in dictionary
                vComps = Split(sDict, "//")
                On Error Resume Next
                    .Range("L" & i).Value = vComps(2)
                    .Range("O" & i).Value = vComps(1)
                    .Range("R" & i).Value = vComps(0)
                On Error GoTo errhandler
            End If
            If i <= 30 Then
                sDict = oDict(.Range("AA" & i).Value)
                If sDict <> "" Then 'doc has value in dictionary
                    vComps = Split(sDict, "//")
                    On Error Resume Next
                        .Range("AC" & i).Value = vComps(0)
                        .Range("AD" & i).Value = vComps(1)
                    On Error GoTo errhandler
                End If
            End If
        Next
    End With
    
    UpdateDocs = True
    Exit Function
errhandler:
    MsgBox "Error updating Document info for " & wsCard.Name & vbCrLf & vbCrLf & "Program terminated"
End Function

Function RearrangeCards(ByRef oDictENG As Object) As Boolean
'''Renames all tabs, then alphabetize (Format: "B-PANDA-789", "C-XIAMEN-988", etc)
'''When a customer's projects are (2), (3), ... --> (1), (2), ... [abandoned]
    Dim i As Integer, j As Integer, k As Integer, iLen As Integer
    Dim sSheetName As String, sCO As String
    Dim sPL As String, sSN As String, sCust As String
    Dim bFound As Boolean, vWS As Variant
    
    iLen = 12 'max length of tab label (without index #)
    
    For i = 4 To ThisWorkbook.Worksheets.Count 'rename if necessary
        sCO = GetCO(ThisWorkbook.Worksheets(i))
        If oDictENG.Exists(sCO) Then '1st letter of PL - Cust(abbr) - SN(short)
            sPL = oDictENG(sCO).Item("PL")
            sCust = oDictENG(sCO).Item("Cust")
            sSN = oDictENG(sCO).Item("SN")
        End If
        If sPL = "" Or sCust = "" Or sSN = "" Then 'something wasn't in Dict
            With ThisWorkbook.Worksheets(i)
                If sCust = "" Then sCust = .Range("C2").Value
                If sSN = "" Then sSN = UCase(.Range("C3").Value)
                If sPL = "" Then
                    If UCase(.Range("C3").Value) Like "C*" Then 'carr
                        sPL = "CARR"
                    ElseIf .Range("C3").Value Like "#####" Then 'mateer
                        sPL = "MATEER"
                    ElseIf .Range("C").Value Like "-###" Then 'burt
                        sPL = "BURT"
                    End If
                End If
            End With
        End If
        If sPL <> "" Then sSheetName = Left(sPL, 1)
        If sSN <> "" Then sSN = ShortenSN(sSN)
        If sCust = "" Then sCust = "???"
        sSheetName = sSheetName & "-" & sCust
        j = iLen - (Len(sSN) + 1)
        If Len(sSheetName) > j Then
            sCust = Replace(sCust, " ", "")
            sSheetName = Left(sPL, 1) & "-" & sCust
            If Len(sSheetName) > j Then sSheetName = Left(sSheetName, j)
        End If
        sSheetName = UCase(sSheetName & "-" & sSN)
        
        For j = 4 To ThisWorkbook.Worksheets.Count 'add # if necessary
            If UCase(ThisWorkbook.Worksheets(j).Name) = sSheetName And i <> j Then
                bFound = True
                k = 1
                Exit For
            End If
        Next
        Do While bFound
            bFound = False
            If k = 1 Then
                sSheetName = sSheetName & "(" & CStr(k) & ")"
            Else
                sSheetName = Left(sSheetName, Len(sSheetName) - Len(CStr(k - 1)) - 1) & CStr(k) & ")"
            End If
            k = k + 1
            For Each vWS In ThisWorkbook.Worksheets
                If UCase(vWS.Name) = sSheetName Then
                    bFound = True
                    Exit For
                End If
            Next
            If k > 100 Then 'something wrong
                sSheetName = "???" & Format(Now, "HH,MM,SS") & CStr(Rnd)
            End If
        Loop
        Debug.Print sSheetName
        ThisWorkbook.Worksheets(i).Name = sSheetName
    Next
    
'shouldn't need numbers on tabs ever... (SNs should be unique)
'    For i = 4 To ThisWorkbook.Worksheets.Count 'renumber as nec. (mult projs for same cust)
'        sSheetName = ThisWorkbook.Worksheets(i).Name
'        j = InStr(sSheetName, "-")
'        If j > 0 Then
'            k = j + 1
'            Do While k < Len(sSheetName)
'                If Mid(sSheetName, k, 1) = "-" Then Exit Do
'                k = k + 1
'            Loop
'            sCust = Left(sSheetName, k) 'including PL because only impt if same PL
'
'        End If
'    Next
    
    k = ThisWorkbook.Worksheets.Count 'alphabetize
    For i = 2 To k - 1
        ThisWorkbook.Worksheets(i).Name = UCase(ThisWorkbook.Worksheets(i).Name)
        For j = i + 1 To k
            If ThisWorkbook.Worksheets(j).Name < ThisWorkbook.Worksheets(i).Name Then
                ThisWorkbook.Worksheets(j).Move Before:=ThisWorkbook.Worksheets(i)
            End If
        Next j
    Next i
    
End Function

Function ShortenSN(sSN As String, Optional iChars As Integer = 9999) As String
'''Shortens SN to designated num of chars (for scorecard tab name)

    ShortenSN = sSN
    
    On Error GoTo errhandler
    
    Dim lSN As Long, i As Integer, j As Integer
    
    sSN = UCase(sSN)
    
    If sSN Like "C*" Then 'carr
        i = 1
        Do While Not IsNumeric(Mid(sSN, i, 1))
            i = i + 1
        Loop
        j = i
        Do While IsNumeric(Mid(sSN, j, 1))
            j = j + 1
        Loop
        lSN = Mid(sSN, i, j - i)
    ElseIf sSN Like "W*" Then 'mateer or burt
        If sSN Like "*-#####*" Then 'mateer
            i = 1
            Do While Not Mid(sSN, i, 5) Like "#####"
                i = i + 1
            Loop
            lSN = Mid(sSN, i, 5)
        ElseIf sSN Like "*-###*" Then 'burt
            i = InStrRev(sSN, "-")
            lSN = Right(sSN, Len(sSN) - i)
        End If
    ElseIf sSN Like "#####" Then 'mateer
        lSN = sSN
    ElseIf sSN Like "###" Then 'carr or burt -- unlikely to be this format
        lSN = sSN
    End If
    
    ShortenSN = CStr(lSN)
    
errhandler:
    If Len(ShortenSN) > iChars Then ShortenSN = Left(ShortenSN, iChars)
End Function
