VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   985
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctDisplay 
      BackColor       =   &H00D0D0D0&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   14355
      TabIndex        =   0
      Top             =   840
      Width           =   14415
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moTime As New clsSevenSegment

Private Enum Mode
    mTime
    mDate
    mCalculator
End Enum

Private Enum Operation
    opNone
    opAdd
    opSub
    opMul
    opDiv
End Enum


Private mmCurrentMode As Mode
Private msAccumulator As String
Private msValue As String
Private msStack As String
Private mopOperation As Operation

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
'Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long

Private Const STRETCH_ANDSCANS = 1
Private Const STRETCH_ORSCANS = 2
Private Const STRETCH_DELETESCANS = 3
Private Const STRETCH_HALFTONE = 4

Private Sub Form_Activate()
    moTime.Initialise pctDisplay, 12, &HD0D0D0, &H202020, &HC0C0C0, 1, 1, 1, 20, 12, 42, 2, 25
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 109 Then
        mmCurrentMode = (mmCurrentMode + 1) Mod 3
    End If
    
    Select Case mmCurrentMode
        Case Mode.mTime
            Timer1.Enabled = True
        Case Mode.mDate
            Timer1.Enabled = True
        Case Mode.mCalculator
            Timer1.Enabled = False
            Select Case KeyAscii
                Case 43
                    mopOperation = opAdd
                Case 45
                    mopOperation = opSub
                Case 42
                    mopOperation = opMul
                Case 47
                    mopOperation = opDiv
                Case Else
                    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
                        If mopOperation = opNone Then
                        Else
                            msStack = msValue
                            msValue = "0"
                        End If
                        msValue = msValue & Chr$(KeyAscii)
                        
                    End If
                    If InStr(msValue, ".") = 0 Then
                        moTime.DisplayFigures msValue & "."
                    Else
                        moTime.DisplayFigures msValue
                    End If
                    
                    Me.Refresh
                    mopOperation = opNone
            End Select
    End Select
End Sub

Private Sub Timer1_Timer()
    Static lTimer As Long
    
    Select Case mmCurrentMode
        Case Mode.mTime
            moTime.DisplayFigures Time
        Case Mode.mDate
            moTime.DisplayFigures Format$(Date, "DD MM YY")
    End Select
    
    Me.Refresh
    SetStretchBltMode Me.hdc, STRETCH_HALFTONE
    StretchBlt Me.hdc, 0, 0, pctDisplay.Width \ 6, pctDisplay.Height \ 6, pctDisplay.hdc, 0, 0, pctDisplay.Width, pctDisplay.Height, &HCC0020
    lTimer = lTimer + 1
End Sub
