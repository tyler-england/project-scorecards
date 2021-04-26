VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Dim wsCard As Worksheet, i As Integer, iRow As Integer
    Dim sPL As String, sCust As String, sMach As String, dDate As Date, sDate As String
    
    Application.ScreenUpdating = False
    Range("A2:Z100").ClearContents
    iRow = 2
    For i = 4 To ThisWorkbook.Worksheets.Count
        Set wsCard = ThisWorkbook.Worksheets(i)
        With wsCard
            sPL = UCase(Left(.Name, 1)) 'product line
            If sPL = "B" Then
                sPL = "Burt"
            ElseIf sPL = "C" Then
                sPL = "Carr"
            ElseIf sPL = "M" Then
                sPL = "Mateer"
            Else
                sPL = "???"
            End If
            sCust = .Range("C2").Value 'customer
            If sCust = "" Then sCust = "???"
            sMach = .Range("C3").Value 'machine / SN
            If sMach = "" Then sMach = "???"
            dDate = .Range("F2").Value
            If dDate > 0 Then
                sDate = Format(dDate, "dd-mmm-yyyy")
            Else
                sDate = "???"
            End If
        End With
        With ThisWorkbook.Worksheets(1)
            .Range("A" & iRow).Value = sPL
            .Range("B" & iRow).Value = sCust
            .Range("C" & iRow).Value = sMach
            .Range("D" & iRow).Value = sDate
        End With
        iRow = iRow + 1
    Next
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(1).Range("I2").Select
End Sub

