VERSION 5.00
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   5655
      Left            =   14040
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
      ScaleWidth      =   925
      TabIndex        =   0
      Top             =   0
      Width           =   13935
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    ShowCalendar pctCalendar, Now, 28, 15, 30, 30
End Sub

Private Sub ShowCalendar(oControl As Control, dStartDate As Date, lWidth As Long, lHeight As Long, lColumnWidth As Long, lRowHeight As Long)
    Dim lDay As Long
    Dim dThisDay As Date
    Dim lColumn As Long
    Dim lRow As Long
    Dim vBackColours As Variant
    Dim lHeader As Long
    Dim lThisDay As Long
    Dim lOffset As Long
    
    vBackColours = Array(vbWhite, &HF0F0F0)
    
    lOffset = 30
    For lHeader = 0 To lWidth - 1
    Next
    
    For lDay = 0 To lWidth * lHeight - 1
        dThisDay = dStartDate + CDate(lDay)
        lColumn = lDay Mod lWidth
        lRow = lDay \ lWidth
        oControl.Line (lColumn * lColumnWidth + lOffset, lRow * lRowHeight)-Step(lColumnWidth, lRowHeight), vBackColours(Month(dThisDay) Mod 2), BF
        lThisDay = Day(dThisDay)
        oControl.CurrentX = lColumn * lColumnWidth + (lColumnWidth - oControl.TextWidth(lThisDay)) \ 2 + lOffset
        oControl.CurrentY = lRow * lRowHeight + oControl.TextHeight(lThisDay) \ 2
        If Weekday(dThisDay) = vbSunday Or Weekday(dThisDay) = vbSaturday Then
            oControl.ForeColor = vbRed
        Else
            oControl.ForeColor = vbBlack
        End If
        
        oControl.Print lThisDay
        If lDay Mod lWidth = lWidth - 1 Then
            If lThisDay <= lWidth Then
                oControl.CurrentX = lWidth * lColumnWidth + lColumnWidth \ 4 + lOffset
                oControl.CurrentY = lRow * lRowHeight + oControl.TextHeight(lThisDay) \ 2
                oControl.Font.Bold = True
                oControl.Print Format$(dThisDay, "MMM")
                oControl.Font.Bold = False
            End If
        ElseIf lDay Mod lWidth = 0 Then
                oControl.CurrentX = 0
                oControl.CurrentY = lRow * lRowHeight + oControl.TextHeight(lThisDay) \ 2
                oControl.Font.Bold = True
                oControl.Print Format$(dThisDay, "MMM")
                oControl.Font.Bold = False
        End If
    Next
End Sub

Private Sub ShowDay(oControl As Control, dStartDate As Date, lWidth As Long, lHeight As Long, lColumnWidth As Long, lRowHeight As Long, dDate As Date)
    Dim lThisDay As Long
    
    lThisDay = Int(dDate) - Int(dStartDate)
    
    oControl.CurrentX = lColumn * lColumnWidth + (lColumnWidth - oControl.TextWidth(lThisDay)) \ 2
    oControl.CurrentY = lRow * lRowHeight + oControl.TextHeight(lThisDay) \ 2
    If Weekday(dThisDay) = vbSunday Or Weekday(dThisDay) = vbSaturday Then
        oControl.ForeColor = vbRed
    Else
        oControl.ForeColor = vbBlack
    End If
    
    oControl.Print lThisDay
End Sub

Private Sub Form_Resize()
    
End Sub
