Attribute VB_Name = "FuncsCards"
Option Explicit

Function CreateNewCards(ByRef sNewCOs() As String) As Boolean
'''Makes new cards & names them appropriately
'''Error returns False
    Dim sCO As String, sSheetName As String
    Dim i As Integer, j As Integer
    Dim bFound As Boolean
    Dim wsBlank As Worksheet, vWS As Variant
    
    On Error GoTo errhandler
    
    Set wsBlank = ThisWorkbook.Worksheets("_BLANK")
    
    For i = 0 To UBound(sNewCOs)
        sCO = sNewCOs(i)
        sSheetName = sCO
        j = 1
        
        For Each vWS In ThisWorkbook.Worksheets
            If vs.Name = sSheetName Then bFound = True
        Next
        
        Do While bFound
            sSheetName = sCO & "-" & CStr(j)
            j = i + 1
            For Each vWS In ThisWorkbook.Worksheets
                If vs.Name = sSheetName Then bFound = True
            Next
            If j > 100 Then 'something wrong
                sshetname = "???" & Format(Now, "HH,MM,SS")
            End If
        Loop
        
        wsBlank.Copy
        With ActiveSheet
            .Name = sSheetName
            .Range("C4").Value = sCO
        End With
    Next
    CreateNewCards = True
errhandler:
End Function

Function UpdateCards(ByRef sCOs() As String, ByRef oDictENG As Object, ByRef oDictActs As Object, ByRef oDictDocs As Object) As Boolean
'''Updates cards for given COs
'''Error returns False
    Dim sCO As String
End Function

