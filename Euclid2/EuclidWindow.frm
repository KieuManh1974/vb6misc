VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   Caption         =   "Euclid"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCanvas As Canvas
Private moLastPosition As Vector
Private moThisPosition As Vector

Private Sub Form_Load()
    Dim oRuler As Ruler
    
    Set moLastPosition = New Vector
    Set moThisPosition = New Vector
    
    Me.ScaleLeft = -Me.ScaleWidth / 2
    Me.ScaleTop = Me.ScaleHeight / 2
    Me.ScaleHeight = -Me.ScaleHeight
    
    Set moCanvas = New Canvas
    Set moCanvas.Surface = Me

    Set oRuler = New Ruler

    moCanvas.AddShape oRuler
    moCanvas.RenderAll
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oDiff As Vector
    
    moThisPosition.Create X, Y
    
    If Button <> vbLeftButton Then
        moCanvas.SelectShape moThisPosition
    Else
        moCanvas.MoveShape moThisPosition.Subtract(moLastPosition), Shift
        moCanvas.RenderAll
    End If
    
    moLastPosition.Create X, Y
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moThisPosition.Create X, Y
    
    moCanvas.DragShape moThisPosition
    
    moLastPosition.Create X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moThisPosition.Create X, Y
    
    moCanvas.ReleaseShapes
   
    moLastPosition.Create X, Y
End Sub
