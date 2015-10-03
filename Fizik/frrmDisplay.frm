VERSION 5.00
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   1680
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moMass As New IShape

Dim moShapes(1) As IShape

Private Sub Form_Activate()
    Dim oRender As IRender
    
    Set moShapes(0) = New clsMass
    Set moShapes(1) = New clsFix
    
    With moShapes(0).InfoList.Member(0)
        .Position.SetComponents 100, 100
        .Velocity.SetComponents 0, -1
        .Force.SetComponents 0, 0.05
        .Mass = 10
    End With
    
    With moShapes(1).InfoList.Member(0)
        .Position.SetComponents 150, 100
    End With
    
    Dim lIndex As Long
    
    For lIndex = LBound(moShapes) To UBound(moShapes)
        Set oRender = moShapes(lIndex)
        Set oRender.DisplayRef = Me
    Next

    tmrTimer.Enabled = True
End Sub


Private Sub Animate()
    Dim lIndex As Long
    Dim oRender As IRender
    
    For lIndex = LBound(moShapes) To UBound(moShapes)
        Set oRender = moShapes(lIndex)
        oRender.Render
        
        moShapes(0).UpdateMotion 0
        DoEvents
    Next
End Sub

Private Sub tmrTimer_Timer()
    Cls
    Animate
End Sub
