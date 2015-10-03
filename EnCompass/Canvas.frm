VERSION 5.00
Begin VB.Form frmCanvas 
   AutoRedraw      =   -1  'True
   Caption         =   "EnCompass"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moGraphicControl As New clsGraphicControl
Private moMousePos As New clsCoordinatePair

Private Sub Form_Initialize()
    moMousePos.SetCoords -1, -1
    Set moGraphicControl.Canvas = Me
    
    moGraphicControl.Test
End Sub

Private Sub Form_Load()
    Form_Resize

    moGraphicControl.Test
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moMousePos.SetCoords X, Y
    moGraphicControl.MouseAction MOUSE_DOWN, moMousePos, Shift And &H1, Shift And &H2, Shift And &H4
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moMousePos.SetCoords X, Y
    moGraphicControl.MouseAction MOUSE_MOVE, moMousePos, Shift And &H1, Shift And &H2, Shift And &H4
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moMousePos.SetCoords X, Y
    moGraphicControl.MouseAction MOUSE_UP, moMousePos, Shift And &H1, Shift And &H2, Shift And &H4
End Sub

Private Sub Form_Resize()
    Dim oNewOffset As New clsCoordinatePair
    
    oNewOffset.SetCoords Me.Width \ Screen.TwipsPerPixelX \ 2, Me.Height \ Screen.TwipsPerPixelY \ 2
    moGraphicControl.Action ACTION_MOVE_OFFSET, oNewOffset
End Sub
