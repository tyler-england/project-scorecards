Attribute VB_Name = "CardUpdate"
Option Explicit

Sub UpdateCurrent()
    Dim sJob As String
    sJob = GetCO(ActiveSheet)
    If sJob = "" Then sJob = InputBox("Enter the CO/job number")
    If sJob <> "" Then
        Call Update(sJob)
    Else
        MsgBox "Unable to update this sheet"
    End If
End Sub

Sub UpdateAll()
    Call Update
End Sub

Function Update(Optional sJob As String = "ALL")
'''updates all scorecards if sJob is omitted

    Dim sCOs() As String, sNewCOs() As String, sUpdatedCOs() As String
    Dim sCO As String
    Dim oDictENG As Object, oDictActs As Object, oDictDocs As Object
    Dim bNew As Boolean, bContinue As Boolean
    Dim i As Integer, j As Integer, x As Integer
    Dim vKeys As Variant
    
    ReDim sCOs(0)
    If UCase(sJob) <> "ALL" Then
        sCOs(0) = sJob
    Else ' (if all)
        With ThisWorkbook
            i = 0
            For x = 1 To .Worksheets.Count
                sCO = GetCO(.Worksheets(x))
                If sCO <> "" Then
                    ReDim Preserve sCOs(i)
                    sCOs(i) = sCO
                    i = i + 1
                End If
            Next
        End With
    End If
    
    Set oDictENG = GetFromENG(sCOs) 'get project data from ENG/ASY (and update ENG/ASY with odictold?)
    If oDictENG Is Nothing Then Exit Function
    
    If UCase(sJob) = "ALL" Then 'find new projects (in ENG wb but not in scorecards yet)
        x = 0
        vKeys = oDictENG.keys
        For i = 0 To oDictENG.Count - 1
            bNew = True
            For j = 0 To UBound(sCOs)
                If UCase(vKeys(i)) = UCase(sCOs(j)) Then
                    bNew = False
                    Exit For
                End If
            Next
            If bNew Then
                ReDim Preserve sNewCOs(x)
                sNewCOs(x) = UCase(vKeys(i))
                x = x + 1
            End If
        Next
        If x > 0 Then 'sNewCOs not empty
            i = UBound(sCOs) + 1
            For x = 0 To UBound(sNewCOs)
                ReDim Preserve sCOs(i + x)
                sCOs(i + x) = sNewCOs(x)
            Next
        End If
    End If
    
    Set oDictActs = GetActuals(sCOs) 'get actual data from CO_Data
    If oDictActs Is Nothing Then Exit Function
    
    Set oDictDocs = GetDocs(sCOs) 'get doc info from doc tracker
    If oDictDocs Is Nothing Then Exit Function

    If x > 0 Then 'sNewCos not empty
        bContinue = CreateNewCards(sNewCOs) 'rename existing & add all new
        If Not bContinue Then Exit Function
    End If
    
    bContinue = UpdateCards(sCOs, oDictENG, oDictActs, oDictDocs) 'update cards
    If Not bContinue Then Exit Function
    
    If Weekday(Date) = 3 Then 'update ENG/MFG workbook
        sUpdatedCOs = UpdateEngMfgApps
        If sUpdatedCOs(0) <> "" Then Call EmailEngMfgApps 'email stakeholders
    End If
    
End Function

Function GetCO(wsCard As Worksheet) As String
'''Returns CO (or SN if no CO) for given worksheet
'''Error returns empty string
    On Error GoTo errhandler
    GetCO = ""
    If wsCard.Range("C4").Value Like "*######*" Then
        If wsCard.Range("C4").Value Like "######" Then
            GetCO = wsCard.Range("C4").Value
        Else 'needs to be parsed (might contain "CO" or something)
            GetCO = wsCard.Range("C4").Value 'TODO: fix this if it's ever an issue
        End If
    ElseIf wsCard.Range("C3").Value Like "C####???" Then 'no CO --> use mach number
        GetCO = Mid(wsCard.Range("C3").Value, 2, 7)
        If Left(GetCO, 1) = "0" Then GetCO = Right(GetCO, Len(GetCO) - 1)
    End If
    Exit Function
errhandler:
    MsgBox "Error getting CO on the following sheet:" & vbCrLf & vbCrLf & _
            wsCard.Name & vbCrLf & vbCrLf & "No updates were made", vbExclamation
End Function

Function EmailEngMfgApps()
'''Sends email to confirm that ENG-MFG-APPS wkbk was updated
    Dim colEmails As New Collection
    
    colEmails.Add "tyler.england@bwpackagingsystems.com"
    
    
    
    Exit Function
errhandler:
    MsgBox "Unable to send email about ENG/MFG/APPS workbook being updated"
End Function
