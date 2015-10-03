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
    
'    oDiagram.AddBox v.Create(10, 10), v.Create(25, 25), &HCCFFFF
'    oDiagram.AddRoundBox v.Create(10, 10), v.Create(25, 25), 5, &HC27C01
'    oDiagram.AddGridCircles v.Create(17.25, 17.25), v.Create(4, 4), 1, v.Create(3, 3), &HCCFFFF
    
'    oDiagram.AddFilledCircle v.Create(30, 30), 15, vbRed
'    oDiagram.AddArrowTriangle v.Create(30, 30), v.Create(20, 20), 0.5, vbWhite
    
    'oDiagram.AddBox v.Create(0, 0), v.Create(60, 60), &HC2C701
'    oDiagram.AddBox v.Create(37 - 11, 30 - 10), v.Create(20, 20), &HC27C01
'    oDiagram.AddChevron v.Create(40, 30), v.Create(10, 12), 4, 0, vbWhite
'    oDiagram.AddChevron v.Create(34, 30), v.Create(10, 12), 4, 0, vbWhite
'
'    oDiagram.AddBox v.Create(13 - 9, 30 - 10), v.Create(20, 20), &HC27C01
'    oDiagram.AddChevron v.Create(10, 30), v.Create(10, 12), 4, 0.5, vbWhite
'    oDiagram.AddChevron v.Create(16, 30), v.Create(10, 12), 4, 0.5, vbWhite

    oDiagram.AddLine v.Create(16, 16), v.Create(8, 8), 2, vbYellow
    oDiagram.AddLine v.Create(16, 16), v.Create(24, 8), 2, vbYellow
'    oDiagram.AddFilledCircle v.Create(10, 22), 3, vbYellow
'    oDiagram.AddFilledCircle v.Create(22, 22), 3, vbYellow
    
    oDiagram.AddBox v.Create(8, 8), v.Create(24, 24), RGB(128, 0, 0)
    oDiagram.AddFilledCircle v.Create(16, 16), 9, vbYellow
    oDiagram.AddFilledCircle v.Create(16, 16), 7.5, vbBlack
    oDiagram.AddFilledCircle v.Create(16, 16), 7, vbWhite
    oDiagram.AddFilledCircle v.Create(16, 16), 1, vbBlack
    oDiagram.AddLine v.Create(16, 16), v.Create(16 + 5, 16), 0.25, vbBlack
    oDiagram.AddLine v.Create(16, 16), v.Create(16, 16 + 6), 0.25, vbBlack
    
    oDiagram.Show
    oDiagram.Render
End Sub
