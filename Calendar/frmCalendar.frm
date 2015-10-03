VERSION 5.00
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   653
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReminder 
      Height          =   5655
      Left            =   6000
      TabIndex        =   1
      Top             =   0
      Width           =   3735
   End
   Begin VB.PictureBox pctCalendar 
      AutoRedraw      =   -1  'True
      Height          =   5655
      Left            =   0
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mnPicWidth As Single
Private mnTextLeft As Single

Private mlColumnWidth As Long
Private mlRowHeight As Long
Private mlDaysWidth As Long
Private mlDaysHeight As Long

Private Sub Form_Load()
    mlColumnWidth = 30
    mlRowHeight = 30
    mnPicWidth = Me.ScaleWidth - pctCalendar.Width
    mnTextLeft = Me.ScaleWidth - txtReminder.Left
End Sub



Private Sub ShowCalendar(oControl As Control, dStartDate As Date, lWidth As Long, lHeight As Long)
    Dim lDay As Long
    Dim dThisDay As Date
    Dim lColumn As Long
    Dim lRow As Long
    Dim vBackColours As Variant
    Dim lHeader As Long
    Dim lThisDay As Long
    Dim lOffset As Long
    
    oControl.Cls
    vBackColours = Array(vbWhite, &HF0F0F0)
    
    lOffset = 30
    For lHeader = 0 To lWidth - 1
    Next
    
    For lDay = 0 To lWidth * lHeight - 1
        dThisDay = dStartDate + CDate(lDay)
        lColumn = lDay Mod lWidth
        lRow = lDay \ lWidth
        oControl.Line (lColumn * mlColumnWidth + lOffset, lRow * mlRowHeight)-Step(mlColumnWidth, mlRowHeight), vBackColours(Month(dThisDay) Mod 2), BF
        lThisDay = Day(dThisDay)
        oControl.CurrentX = lColumn * mlColumnWidth + (mlColumnWidth - oControl.TextWidth(lThisDay)) \ 2 + lOffset
        oControl.CurrentY = lRow * mlRowHeight + oControl.TextHeight(lThisDay) \ 2
        If Weekday(dThisDay) = vbSunday Or Weekday(dThisDay) = vbSaturday Then
            oControl.ForeColor = vbRed
        Else
            oControl.ForeColor = vbBlack
        End If
        
        oControl.Print lThisDay
        oControl.ForeColor = vbBlack
        If lDay Mod lWidth = lWidth - 1 Then
            If lThisDay <= lWidth Then
                oControl.CurrentX = lWidth * mlColumnWidth + mlColumnWidth \ 4 + lOffset
                oControl.CurrentY = lRow * mlRowHeight + oControl.TextHeight(lThisDay) \ 2
                oControl.Font.Bold = True
                oControl.Print Format$(dThisDay, "MMM")
                oControl.Font.Bold = False
            End If
        ElseIf lDay Mod lWidth = 0 Then
                oControl.CurrentX = 0
                oControl.CurrentY = lRow * mlRowHeight + oControl.TextHeight(lThisDay) \ 2
                oControl.Font.Bold = True
                oControl.Print Format$(dThisDay, "MMM")
                oControl.Font.Bold = False
        End If
    Next
End Sub

Private Sub ShowDay(oControl As Control, dStartDate As Date, lColumn As Long, lRow As Long, Optional lBorderColour As Long = -1)
    Dim lThisDay As Long
    Dim dThisDay As Date
    
    lThisDay = lColumn + lRow * mlDaysWidth
    dThisDay = dStartDate + CDate(lThisDay)
    
    oControl.Line (lColumn * mlColumnWidth + 30, lRow * mlRowHeight)-Step(mlColumnWidth - 1, mlRowHeight - 1), lBorderColour, B
    
    oControl.CurrentX = lColumn * mlColumnWidth + (mlColumnWidth - oControl.TextWidth(lThisDay)) \ 2 + 30
    oControl.CurrentY = lRow * mlRowHeight + oControl.TextHeight(lThisDay) \ 2
    If Weekday(dThisDay) = vbSunday Or Weekday(dThisDay) = vbSaturday Then
        oControl.ForeColor = vbRed
    Else
        oControl.ForeColor = vbBlack
    End If
    
    oControl.Print Day(dThisDay)
End Sub


Private Sub Form_Resize()
    pctCalendar.Width = Me.ScaleWidth - mnPicWidth
    txtReminder.Left = Me.ScaleWidth - mnTextLeft
    pctCalendar.Height = Me.Height
    txtReminder.Height = Me.Height

    mlDaysWidth = ((pctCalendar.ScaleWidth \ mlColumnWidth - 2) \ 7) * 7
    mlDaysHeight = pctCalendar.ScaleHeight \ mlRowHeight
    
    ShowCalendar pctCalendar, Now, mlDaysWidth, mlDaysHeight
End Sub


Private Sub pctCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lColumn As Long
    Dim lRow As Long
    
    If (X - 30) < 0 Then
        Exit Sub
    End If
    lColumn = Int((X - 30) / mlColumnWidth)
    lRow = Y \ mlRowHeight
    
    If lColumn >= 0 And lColumn <= mlDaysWidth And lRow >= 0 And lRow <= mlDaysHeight Then
        ShowDay pctCalendar, Now, lColumn, lRow, vbRed
    End If
End Sub
