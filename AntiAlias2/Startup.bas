Attribute VB_Name = "Startup"
Option Explicit

Public Sub Main()
    Dim oDiagram As New clsFineDiagram
    Dim v As New clsVector
    
    'oDiagram.AddLine v.Create(10, 10), v.Create(100, 100), 32 / 16, vbRed
    'oDiagram.AddCircle v.Create(70, 70), 50, 1, vbGreen
    'oDiagram.AddFilledSector v.Create(70, 70), 50, 0.5, 2, vbRed
    'oDiagram.AddSector v.Create(70, 70), 50, 10, 0.5, 2, vbRed
    'oDiagram.AddFilledTriangle v.Create(10, 10), v.Create(40, 20), v.Create(30, 50), vbRed
    
    oDiagram.AddBox v.Create(10, 10), v.Create(25, 25), &HCCFFFF
    oDiagram.AddRoundBox v.Create(10, 10), v.Create(25, 25), 5, &HC27C01
    oDiagram.AddGridCircles v.Create(17.25, 17.25), v.Create(5, 5), 1.5, v.Create(2, 2), &HCCFFFF
    oDiagram.Show
    oDiagram.Render
End Sub
