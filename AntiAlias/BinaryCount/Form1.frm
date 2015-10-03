VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const hw = 16
Const hseparation = 2 * hw + 24
Const vseparation = 2 * hw + 40
Const lmargin = 30
Const tmargin = 30

Private Sub Form_Load()
    DoEvents
    DoEvents
    ShowTiles
End Sub

Private Function ShowTiles()
    Dim iIndex As Long
    Dim xpos As Single
    Dim ypos As Single
    
    Printer.ScaleMode = 3
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 4
    Printer.ForeColor = vbBlue
    
    DrawWidth = 1
    For iIndex = 0 To 255
        xpos = (iIndex Mod 16) * hseparation + lmargin
        ypos = (iIndex \ 16) * vseparation + tmargin
        
        ShowGraph iIndex, xpos, ypos
        Line (xpos + hw - TextWidth(CStr(iIndex)) / 2, ypos + hw * 2 + 4)-Step(0, 0)
        Printer.Line (xpos + hw - Printer.TextWidth(CStr(iIndex)) / 2, ypos + hw * 2 + 4)-Step(0, 0)
        
        Print ; iIndex
        Printer.Print ; iIndex
    Next
    Printer.EndDoc
End Function

Private Function ShowGraph(ByVal number As Long, ByVal x As Single, ByVal y As Single)
    Dim vPositions As Variant
    Dim iIndex As Long
    Dim sBin As String
    Dim vPosition As Variant
    
    vPositions = Array(Array(2, 2, 0, -2), Array(2, 0, -2, 0), Array(0, 0, 0, 2), Array(0, 2, 2, 0), Array(0, 2, 2, -2), Array(1, 2, 0, -2), Array(2, 2, -2, -2), Array(2, 1, -2, 0))
    sBin = Binary(number)
    
    For iIndex = 8 To 1 Step -1
        vPosition = vPositions(8 - iIndex)
        'Line (x + (vPosition(0) + vPosition(2)) * hw / 2, y + (vPosition(1) + vPosition(3)) * hw / 2)-Step(0, 0)
        'Printer.Line (x + (vPosition(0) + vPosition(2)) * hw / 2, y + (vPosition(1) + vPosition(3)) * hw / 2)-Step(0, 0)
        Line (x + hw, y + hw)-Step(1, 1)
        Printer.Line (x + hw, y + hw)-Step(1, 1)
        If Mid$(sBin, iIndex, 1) = "1" Then
            Line (x + vPosition(0) * hw, y + vPosition(1) * hw)-Step(vPosition(2) * hw, vPosition(3) * hw)
            Printer.Line (x + vPosition(0) * hw, y + vPosition(1) * hw)-Step(vPosition(2) * hw, vPosition(3) * hw)
        End If
    Next
End Function

Private Function Binary(ByVal x As Long) As String
    While x > 0
        Binary = (x Mod 2) & Binary
        x = x \ 2
    Wend
    
    Binary = String$(8 - Len(Binary), "0") & Binary
End Function
