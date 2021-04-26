VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub ChartAxis()
'''updates graph Y-range to get rid of empty space
    Dim siMin As Single, siMax As Single
    Dim iMin As Integer, iMax As Integer
    Dim iTest As Integer, i As Integer
    Dim v As Variant, cChart As ChartObject
    With Me
        siMin = WorksheetFunction.Min(.Range("AJ:AK"))
        siMax = WorksheetFunction.Max(.Range("AJ:AK"))
    End With
    
    If siMin = siMax Then Exit Sub
    
    iMin = siMin * 100
    iMax = siMax * 100
    
    iTest = iMin
    Do While Not iTest < iMin
        iTest = WorksheetFunction.MRound(iMin - i, 5)
        i = i + 1
    Loop
    siMin = iTest / 100
    
    iTest = iMax
    i = 0
    Do While Not iTest > iMax
        iTest = WorksheetFunction.MRound(iMax + i, 5)
        i = i + 1
    Loop
    siMax = iTest / 100
'
'    For Each v In ActiveSheet.ChartObjects
'        If InStr(v.Name, "Chart 1") > 0 Then
'            Set cChart = v
'            Exit For
'        End If
'    Next
'
    With ActiveSheet.ChartObjects(1).Chart.Axes(xlValue, xlPrimary)
        .MinimumScale = siMin
        .MaximumScale = siMax
    End With
End Sub

Private Sub Worksheet_Activate()
    Call ChartAxis
End Sub
