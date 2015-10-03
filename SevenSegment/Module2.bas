Attribute VB_Name = "Module2"
Option Explicit

Private Declare Function PolyPolygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Private Declare Function CreatePenIndirect Lib "gdi32.dll" (lpLogPen As LOGPEN) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
 
Private Type POINTAPI
   x As Long
   y As Long
End Type
 
Private Type LOGPEN
   lopnStyle As Long
   lopnWidth As POINTAPI
   lopnColor As Long
End Type
 
Private Const PS_DASH = 1
Private Const PS_DOT = 2
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4
Private Const PS_NULL = 5
Private Const PS_INSIDEFRAME = 6
Private Const PS_SOLID = 0

Private Sub Form_Load()
   Dim Retval As Long
   Dim hPen As Long
   Dim hOldPen As Long
   Dim PenInfo As LOGPEN
   Dim Ecken(7) As POINTAPI
   Dim Objekte(1) As Long
   
   Me.AutoRedraw = True
   Me.ScaleMode = vbPixels

   With PenInfo
      .lopnColor = vbRed
      .lopnStyle = PS_SOLID
   End With

   hPen = CreatePenIndirect(PenInfo)

   hOldPen = SelectObject(Me.hdc, hPen)

   With Ecken(0)
      .x = Me.ScaleWidth / 2
      .y = 0
   End With
   With Ecken(1)
      .x = Me.ScaleWidth
      .y = Me.ScaleHeight / 4
   End With
   With Ecken(2)
      .x = Me.ScaleWidth / 2
      .y = Me.ScaleHeight / 2
   End With
   With Ecken(3)
      .x = 0
      .y = Me.ScaleHeight / 4
   End With
   Objekte(0) = 4

   With Ecken(4)
      .x = Me.ScaleWidth / 2
      .y = Me.ScaleHeight / 2
   End With
   With Ecken(5)
      .x = Me.ScaleWidth
      .y = Me.ScaleHeight / 4 * 3
   End With
   With Ecken(6)
      .x = Me.ScaleWidth / 2
      .y = Me.ScaleHeight
   End With
   With Ecken(7)
      .x = 0
      .y = Me.ScaleHeight / 4 * 3
   End With
   Objekte(1) = 4

   Retval = PolyPolygon(Me.hdc, Ecken(0), Objekte(0), UBound(Objekte) + 1)

   Call SelectObject(Me.hdc, hOldPen)

   DeleteObject hPen
End Sub
