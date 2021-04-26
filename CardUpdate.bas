Attribute VB_Name = "CardUpdate"
Option Explicit

Sub UpdateCurrent()
    Dim sJob As String, wsCurrent As Worksheet
    Set wsCurrent = ActiveSheet
    sJob = GetCO(ActiveSheet)
    If sJob = "" Then sJob = InputBox("Enter the CO/job number")
    If sJob <> "" Then
        Call Update(sJob)
    Else
        MsgBox "Unable to update this sheet"
    End If
    wsCurrent.Activate
End Sub

Sub UpdateAll()
    Call Update
    ThisWorkbook.Worksheets(1).Activate
End Sub

Function Update(Optional sJob As String = "ALL")
'''updates all scorecards if sJob is omitted

    Dim sCOs() As String, sNewCOs() As String, sUpdatedCOs() As String
    Dim sCO As String
    Dim oDictENG As Object, oDictActs As Object, oDictDocs As Object
    Dim bNew As Boolean, bContinue As Boolean
    Dim i As Integer, j As Integer, x As Integer
    Dim vKeys As Variant
    
    On Error GoTo errhandler
    
    Application.StatusBar = "Getting COs..."
    
    ReDim sCOs(0)
    If UCase(sJob) <> "ALL" Then
        sCOs(0) = sJob
    ElseIf ThisWorkbook.Worksheets.Count > 4 Then ' (if all)
        With ThisWorkbook
            i = 0
            For x = 4 To .Worksheets.Count
                sCO = GetCO(.Worksheets(x))
                If sCO <> "" Then
                    ReDim Preserve sCOs(i)
                    sCOs(i) = sCO
                    i = i + 1
                End If
            Next
        End With
    End If
    
    
    
    Application.StatusBar = "Checking ENG/MFG/APP file..."
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
            If UBound(sCOs) = 0 And sCOs(0) = "" Then i = 0 'no "old" COs
            For x = 0 To UBound(sNewCOs)
                ReDim Preserve sCOs(i + x)
                sCOs(i + x) = sNewCOs(x)
            Next
        End If
    End If
    
    On Error Resume Next
        sCO = sCOs(0)
        If Err.Number <> 0 Then
            MsgBox "No COs"
        End If
    On Error GoTo exitline
    
    Application.StatusBar = "Checking XA info..."
    Set oDictActs = GetActuals(sCOs) 'get actual data from CO_Data
    If oDictActs Is Nothing Then GoTo exitline
    
    Application.StatusBar = "Checking Doc Tracker..."
    Set oDictDocs = GetDocs(sCOs) 'get doc info from doc tracker
    If oDictDocs Is Nothing Then GoTo exitline

    If x > 0 Then 'sNewCos not empty
        Application.StatusBar = "Making cards for new projects..."
        bContinue = CreateNewCards(sNewCOs) 'add all new
        If Not bContinue Then GoTo exitline
    End If
    
    Application.StatusBar = "Updating cards..."
    bContinue = UpdateCards(sCOs, oDictENG, oDictActs, oDictDocs) 'update cards & rename as necessary
    If Not bContinue Then GoTo exitline
    
    If Weekday(Date) = 3 Then 'update ENG/MFG workbook
        sUpdatedCOs = UpdateEngMfgApps(oDictActs)
        If sUpdatedCOs(0) <> "" Then Call EmailEngMfgApps(sUpdatedCOs) 'email stakeholders
    End If
    Application.StatusBar = False
    Exit Function
errhandler:
    MsgBox "Error updating the cards"
exitline:
    Application.StatusBar = False
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
    ElseIf wsCard.Range("C3").Value > 0 Then
        GetCO = wsCard.Range("C3").Value
    Else
        GetCO = "???" & Format(Now, "HHMMSS") & CStr(Rnd)
    End If
    Exit Function
errhandler:
    MsgBox "Error getting CO on the following sheet:" & vbCrLf & vbCrLf & _
            wsCard.Name & vbCrLf & vbCrLf & "No updates were made", vbExclamation
End Function

Function EmailEngMfgApps(sUpdatedCOs() As String)
'''Sends email to confirm that ENG-MFG-APPS wkbk was updated
    Dim oOutlook As Object, oMail As Object
    Dim colEmails As New Collection
    Dim var As Variant
    Dim sName As String
    
    colEmails.Add "tyler.england@bwpackagingsystems.com"
    
    sName = Application.UserName
    On Error Resume Next
        sName = Right(sName, Len(sName) - InStrRev(sName, ","))
        sName = Left(sName, InStr(sName, "(") - 1)
        sName = Trim(sName)
    On Error GoTo errhandler
    
    Set oOutlook = CreateObject("Outlook.Application") 'email Kim & Pam & John
    oOutlook.Session.Logon
    Set oMail = oOutlook.CreateItem(olMailItem)
    With oMail
        For Each var In colEmails
            .Recipients.Add var
        Next
        .Subject = "ENG-MFG-APP Workbook Updated"
        .Body = "Hi all," & vbCrLf & vbCrLf & "The ENG/MFG/APP scorecard workbook has just been updated. " & _
                "The following projects have had hour values updated:" & vbCrLf & vbCrLf & Join(sUpdatedCOs, vbCrLf) & _
                vbCrLf & vbCrLf & "(This is an automated email)" & vbCrLf & vbCrLf & "Best," & vbCrLf & sName
    End With
    oMail.Display
    'oMail.Send
    
    Exit Function
errhandler:
    MsgBox "Unable to send email about ENG/MFG/APPS workbook being updated"
End Function
