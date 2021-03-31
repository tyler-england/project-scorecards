Attribute VB_Name = "Setup"
Option Explicit
Public arrErrorEmails() As String, iNumMsgs As Integer, bNewMsg As Boolean 'for ErrorRep
Public wbDataCO As String, wbDataENG As String, wbDocTracker As String

Function GlobalVars()
    ''''''hardcoded variables''''''''''
    wbDataCO = "https://bw1-my.sharepoint.com/personal/tyler_england_bwpackagingsystems_com/Documents/Distributed Files/Project Scorecards/CO_Data.xlsm"
    wbDataENG = "https://bw1-my.sharepoint.com/personal/tyler_england_bwpackagingsystems_com/Documents/Distributed Files/Project Scorecards/ENG-MFG-APP_Data.xlsm"
    wbDocTracker = "https://bw1-my.sharepoint.com/personal/tyler_england_bwpackagingsystems_com/Documents/Distributed Files/Doc Tracker.xlsm"
End Function

Function ExportModules() As Boolean
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'" & wbMacro.Name & "'!ExportModules", ThisWorkbook
    If Not bOpen Then wbMacro.Close savechanges:=False
    ExportModules = True
End Function

Public Sub ErrorRep(rouName, rouType, curVal, errNum, errDesc, miscInfo)
    
    Dim oApp As Object, oEmail As MailItem, arrEmailTxt(10) As String
    Dim outlookOpen As Boolean, emailTxt As String, varMsg As Variant
    
    Application.ScreenUpdating = False
    arrEmailTxt(2) = "--Issue finding Workbook"
    arrEmailTxt(3) = "--Issue finding User"
    arrEmailTxt(4) = "--Issue finding Workbook path"
    arrEmailTxt(5) = "--Issue finding Routine name"
    arrEmailTxt(6) = "--Issue finding Routine type"
    arrEmailTxt(7) = "--Issue finding Current value"
    arrEmailTxt(8) = "--Issue finding Error number"
    arrEmailTxt(9) = "--Issue finding Error description"
    arrEmailTxt(10) = "--Issue finding Misc. add'l info"
    
    On Error Resume Next
        Set oApp = GetObject(, "Outlook.Application")
        outlookOpen = True
        
        ''''''can't use error handler because these varTypes might be problematic
        If Not VarType(curVal) = vbString Then 'make into string
            If VarType(curVal) > 8000 Then 'array of some sort
                curVal = Join(curVal, ";")
            Else 'hopefully this will make it a string
                curVal = Str(curVal)
            End If
        End If
        
        If Not VarType(miscInfo) = vbString Then 'make into string
            If VarType(miscInfo) > 8000 Then 'array of some sort
                curVal = Join(miscInfo, ";")
            Else 'hopefully this will make it a string
                curVal = Str(miscInfo)
            End If
        End If
        
    On Error Resume Next 'types might cause errors
        arrEmailTxt(0) = "REPORT"
        arrEmailTxt(1) = "Error occurred in VBA program. Details are listed below." & vbCrLf
        arrEmailTxt(2) = Right(arrEmailTxt(2), Len(arrEmailTxt(2)) - 16) & ": " & ThisWorkbook.Name
        arrEmailTxt(3) = Right(arrEmailTxt(3), Len(arrEmailTxt(3)) - 16) & ": " & Application.UserName & vbCrLf
        arrEmailTxt(4) = Right(arrEmailTxt(4), Len(arrEmailTxt(4)) - 16) & ": " & ThisWorkbook.Path
        arrEmailTxt(5) = Right(arrEmailTxt(5), Len(arrEmailTxt(5)) - 16) & ": " & rouName
        arrEmailTxt(6) = Right(arrEmailTxt(6), Len(arrEmailTxt(6)) - 16) & ": " & rouType
        arrEmailTxt(7) = Right(arrEmailTxt(7), Len(arrEmailTxt(7)) - 16) & ": " & curVal & vbCrLf
        arrEmailTxt(8) = Right(arrEmailTxt(8), Len(arrEmailTxt(8)) - 16) & ": " & errNum
        arrEmailTxt(9) = Right(arrEmailTxt(9), Len(arrEmailTxt(9)) - 16) & ": " & errDesc & vbCrLf
        arrEmailTxt(10) = Right(arrEmailTxt(10), Len(arrEmailTxt(10)) - 16) & ": " & vbCrLf & miscInfo
    On Error GoTo errhandler
    
    emailTxt = Join(arrEmailTxt, vbCrLf)
    
    'see if emailTxt has been sent already this session
    bNewMsg = True 'default value
    If iNumMsgs > 0 Then 'at least one email has been generated already
        For Each varMsg In arrErrorEmails 'see if there were any matches
            If UCase(varMsg) = UCase(emailTxt) Then 'this was already sent this session
                bNewMsg = False
                Exit For
            End If
        Next
    End If
    
    If bNewMsg Then 'new message -> add to array for next time
        iNumMsgs = iNumMsgs + 1
        ReDim Preserve arrErrorEmails(iNumMsgs)
        arrErrorEmails(iNumMsgs) = emailTxt
    Else 'repeat message
        Exit Sub
    End If
    
    If oApp Is Nothing Then
        Set oApp = CreateObject("Outlook.Application")
        outlookOpen = False
    End If
    
    Set oEmail = oApp.CreateItem(0)

    With oEmail
        .To = "tyler.england@bwpackagingsystems.com"
        .Subject = "VBA Program Error Report"
        .Body = emailTxt
        If InStr(UCase(Application.UserName), "ENGLAND, TYLER") > 0 Then
            .Display 'it me
        Else:
            .Send
        End If
    End With
    
    If Not outlookOpen Then oApp.Close
errhandler:
End Sub




