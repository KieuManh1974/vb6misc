VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSevenSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Type POINTAPI
   x As Long
   y As Long
End Type


Private mlWidthVert As Long
Private mlWidthHoriz As Long
Private mlHeightVert As Long
Private mlHeightHoriz As Long
Private mlSeparationVert As Long
Private mlSeparationHoriz As Long
Private mvDigitPatterns(37) As Variant
Private Const msDigits As String = " 0123456789ABCEFGHJLOPSUbcdhinortu-.:"
    
Private mlSegmentOnColour As Long
Private mlSegmentOffColour As Long
Private mlBackplaneColour As Long

Private mlScaling As Long
Private mlScalingVert As Long
Private mlScalingHoriz As Long
Private mlWidth As Long
Private mlLength As Long
Private mlSeparation As Long
Private mlDigitSeparation As Long
Private mlPadding

Private mlTotalDigits As Long

Private moForm As Object

Public Sub Initialise(oForm As Object, lTotalDigits As Long, lBackplaneColour As Long, lSegmentOnColour As Long, lSegmentOffColour As Long, lScaling As Long, lScalingVert As Long, lScalingHoriz As Long, lPadding As Long, lWidth As Long, lLength As Long, lSeparation As Long, lDigitSeparation As Long)
    Set moForm = oForm
    mlBackplaneColour = lBackplaneColour
    mlSegmentOnColour = lSegmentOnColour
    mlSegmentOffColour = lSegmentOffColour
    mlScaling = lScaling
    mlScalingVert = lScalingVert
    mlScalingHoriz = lScalingHoriz
    mlPadding = lPadding
    mlWidth = lWidth
    mlLength = lLength
    mlSeparation = lSeparation
    mlDigitSeparation = lDigitSeparation
    mlTotalDigits = lTotalDigits
    
    Activate
End Sub

Private Sub Activate()
    AdjustMetrics
    ShowBackplane
    Reset
End Sub

Private Sub AdjustMetrics()
    mlWidthVert = mlScaling * mlScalingHoriz * mlWidth
    mlWidthHoriz = mlScaling * mlScalingHoriz * mlLength
    mlHeightVert = mlScaling * mlScalingVert * mlLength
    mlHeightHoriz = mlScaling * mlScalingVert * mlWidth
    mlSeparationHoriz = mlScaling * mlScalingHoriz * mlSeparation
    mlSeparationVert = mlScaling * mlScalingVert * mlSeparation
    
    mvDigitPatterns(0) = Array()
    mvDigitPatterns(1) = Array() ' blank
    mvDigitPatterns(2) = Array(1, 2, 3, 5, 6, 7) '0
    mvDigitPatterns(3) = Array(3, 6)
    mvDigitPatterns(4) = Array(1, 3, 4, 5, 7)
    mvDigitPatterns(5) = Array(1, 3, 4, 6, 7)
    mvDigitPatterns(6) = Array(2, 3, 4, 6)
    mvDigitPatterns(7) = Array(1, 2, 4, 6, 7)
    mvDigitPatterns(8) = Array(1, 2, 4, 5, 6, 7)
    mvDigitPatterns(9) = Array(1, 3, 6)
    mvDigitPatterns(10) = Array(1, 2, 3, 4, 5, 6, 7)
    mvDigitPatterns(11) = Array(1, 2, 3, 4, 6, 7) ' 9
    mvDigitPatterns(12) = Array(1, 2, 3, 4, 5, 6) 'A
    mvDigitPatterns(13) = Array(1, 2, 3, 4, 5, 6, 7)
    mvDigitPatterns(14) = Array(1, 2, 5, 7)
    mvDigitPatterns(15) = Array(1, 2, 4, 5, 7)
    mvDigitPatterns(16) = Array(1, 2, 4, 5)
    mvDigitPatterns(17) = Array(1, 2, 5, 6, 7)
    mvDigitPatterns(18) = Array(2, 3, 4, 5, 6)
    mvDigitPatterns(19) = Array(3, 5, 6, 7)
    mvDigitPatterns(20) = Array(2, 5, 7)
    mvDigitPatterns(21) = Array(1, 2, 3, 5, 6, 7)
    mvDigitPatterns(22) = Array(1, 2, 3, 4, 5)
    mvDigitPatterns(23) = Array(1, 2, 4, 6, 7)
    mvDigitPatterns(24) = Array(2, 3, 5, 6, 7) 'U
    mvDigitPatterns(25) = Array(2, 4, 5, 6, 7) ' b
    mvDigitPatterns(26) = Array(4, 5, 7)
    mvDigitPatterns(27) = Array(3, 4, 5, 6, 7)
    mvDigitPatterns(28) = Array(2, 4, 5, 6)
    mvDigitPatterns(29) = Array(6)
    mvDigitPatterns(30) = Array(4, 5, 6)
    mvDigitPatterns(31) = Array(4, 5, 6, 7)
    mvDigitPatterns(32) = Array(4, 5)
    mvDigitPatterns(33) = Array(2, 4, 5, 7)
    mvDigitPatterns(34) = Array(5, 6, 7) 'u
    mvDigitPatterns(35) = Array(4) ' minus
    mvDigitPatterns(36) = Array(8) ' dot
    mvDigitPatterns(37) = Array(9) ' colon
End Sub

Private Sub ShowBackplane()
    Dim hPen As Long
    Dim hBrush As Long
    Dim hOldPen As Long
    Dim hOldBrush As Long
    Dim paPoints(3) As POINTAPI
    Dim lColour As Long
    Dim lX As Long
    Dim lY As Long
    
    lX = 0
    lY = 0
    
    hPen = CreatePen(0, 0, mlBackplaneColour)
    hBrush = CreateSolidBrush(mlBackplaneColour)
    
    hOldPen = SelectObject(moForm.hdc, hPen)
    hOldBrush = SelectObject(moForm.hdc, hBrush)
    
    With paPoints(0)
        .x = lX
        .y = lY
    End With
    With paPoints(1)
        .x = lX + mlTotalDigits * (mlWidthHoriz + mlSeparationHoriz * 3 + mlDigitSeparation * mlScaling) + mlScaling * mlPadding * 2
        .y = lY
    End With
    With paPoints(2)
        .x = lX + mlTotalDigits * (mlWidthHoriz + mlSeparationHoriz * 3 + mlDigitSeparation * mlScaling) + mlScaling * mlPadding * 2
        .y = lY + mlHeightVert * 2 + mlSeparationVert * 3 + mlScaling * mlPadding * 2
    End With
    With paPoints(3)
        .x = lX
        .y = lY + mlHeightVert * 2 + mlSeparationVert * 3 + mlScaling * mlPadding * 2
    End With

    Polygon moForm.hdc, paPoints(0), UBound(paPoints) + 1
    
    Call SelectObject(moForm.hdc, hOldPen)
    Call SelectObject(moForm.hdc, hOldBrush)
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Public Sub DisplayFigures(ByVal sFigure As String)
    Dim lDigitIndex As Long
    Dim lLength As Long
    Dim sDigit As String
    Dim lOffset As Long
    
    lLength = Len(sFigure)
    
    Reset
    
    For lDigitIndex = 0 To lLength - 1
        sDigit = Mid$(sFigure, lDigitIndex + 1, 1)
        If sDigit = "." Then
            lOffset = lOffset + 1
        End If
        DisplayDigit sDigit, mlTotalDigits - lLength + lDigitIndex - lOffset
    Next
End Sub

Private Sub Reset()
    Dim lDigitIndex As Long
    
    For lDigitIndex = 0 To mlTotalDigits - 1
        ResetDigit lDigitIndex
    Next
End Sub

Private Sub ResetDigit(ByVal lPosition As Long)
    Dim lSegmentIndex As Long
    For lSegmentIndex = 1 To 9
        ShowSegment lSegmentIndex, lPosition, False
    Next
End Sub

Private Sub DisplayDigit(ByVal sDigit As String, ByVal lPosition As Long)
    Dim vDigit As Variant
    Dim lSegmentIndex As Long
    Dim lDigit As Long
    Dim lIndex As Long
    
    Dim sDigits As String

    For Each vDigit In mvDigitPatterns(InStr(msDigits, sDigit))
        ShowSegment vDigit, lPosition, True
    Next
End Sub


Private Sub ShowSegment(ByVal lSegmentIndex, ByVal lPosition As Long, ByVal bOn As Boolean)
    Dim lHeight1 As Long
    Dim lHeight2 As Long
    Dim lHeight3 As Long
    Dim lLeft1 As Long
    Dim lLeft2 As Long
    
    lHeight1 = mlPadding * mlScaling
    lHeight2 = lHeight1 + mlHeightVert
    lHeight3 = lHeight2 + mlHeightVert
    
    lLeft1 = mlPadding * mlScaling + lPosition * (mlWidthHoriz + mlSeparationHoriz * 3 + mlDigitSeparation * mlScaling)
    lLeft2 = lLeft1 + mlWidthHoriz
    
    Select Case lSegmentIndex
        Case 1
            DrawHorizontalSegment lLeft1 + mlSeparationHoriz, lHeight1, bOn
        Case 2
            DrawVerticalSegment lLeft1, lHeight1 + mlSeparationVert, bOn
        Case 3
            DrawVerticalSegment lLeft2 + mlSeparationHoriz * 2, lHeight1 + mlSeparationVert, bOn
        Case 4
            DrawHorizontalSegment lLeft1 + mlSeparationHoriz, lHeight2 + mlSeparationVert * 2, bOn
        Case 5
            DrawVerticalSegment lLeft1, lHeight2 + mlSeparationVert * 3, bOn
        Case 6
            DrawVerticalSegment lLeft2 + mlSeparationHoriz * 2, lHeight2 + mlSeparationVert * 3, bOn
        Case 7
            DrawHorizontalSegment lLeft1 + mlSeparationHoriz, lHeight3 + mlSeparationVert * 4, bOn
        Case 8
            DrawDot lLeft2 + mlSeparationHoriz * 3, lHeight3 + mlSeparationVert * 4, bOn
        Case 9
            DrawDot lLeft1 + mlSeparationHoriz * 0 + mlWidthHoriz \ 2, lHeight1 + mlSeparationVert * 1 + mlHeightVert \ 2, bOn
            DrawDot lLeft1 + mlSeparationHoriz * 0 + mlWidthHoriz \ 2, lHeight2 + mlSeparationVert * 1 + mlHeightVert \ 2, bOn
    End Select
End Sub

Private Sub DrawDot(ByVal lX As Long, ByVal lY As Long, ByVal bOn As Boolean)
    Dim hPen As Long
    Dim hBrush As Long
    Dim hOldPen As Long
    Dim hOldBrush As Long
    Dim paPoints(5) As POINTAPI
    Dim lColour As Long
    
    If bOn Then
        lColour = mlSegmentOnColour
    Else
        lColour = mlSegmentOffColour
    End If
    
    hPen = CreatePen(0, 0, lColour)
    hBrush = CreateSolidBrush(lColour)
    
    hOldPen = SelectObject(moForm.hdc, hPen)
    hOldBrush = SelectObject(moForm.hdc, hBrush)
    
    Ellipse moForm.hdc, lX, lY, lX + mlWidthVert * 2 - mlWidthVert \ 2, lY + mlHeightHoriz * 2 - mlHeightHoriz \ 2
    
    Call SelectObject(moForm.hdc, hOldPen)
    Call SelectObject(moForm.hdc, hOldBrush)
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Private Sub DrawVerticalSegment(ByVal lX As Long, ByVal lY As Long, ByVal bOn As Boolean)
    Dim hPen As Long
    Dim hBrush As Long
    Dim hOldPen As Long
    Dim hOldBrush As Long
    Dim paPoints(5) As POINTAPI
    Dim lColour As Long
    
    If bOn Then
        lColour = mlSegmentOnColour
    Else
        lColour = mlSegmentOffColour
    End If
    
    hPen = CreatePen(0, 0, lColour)
    hBrush = CreateSolidBrush(lColour)
    
    hOldPen = SelectObject(moForm.hdc, hPen)
    hOldBrush = SelectObject(moForm.hdc, hBrush)
    
    With paPoints(0)
        .x = lX
        .y = lY
    End With
    With paPoints(1)
        .x = lX + mlWidthVert \ 2
        .y = lY + mlHeightHoriz \ 2
    End With
    With paPoints(2)
        .x = lX + mlWidthVert \ 2
        .y = lY + mlHeightVert - mlHeightHoriz \ 2
    End With
    With paPoints(3)
        .x = lX
        .y = lY + mlHeightVert
    End With
    With paPoints(4)
        .x = lX - mlWidthVert \ 2
        .y = lY + mlHeightVert - mlHeightHoriz \ 2
    End With
    With paPoints(5)
        .x = lX - mlWidthVert \ 2
        .y = lY + mlHeightHoriz \ 2
    End With
    
    Polygon moForm.hdc, paPoints(0), UBound(paPoints) + 1
    
    Call SelectObject(moForm.hdc, hOldPen)
    Call SelectObject(moForm.hdc, hOldBrush)
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Private Sub DrawHorizontalSegment(ByVal lX As Long, ByVal lY As Long, ByVal bOn As Boolean)
    Dim hPen As Long
    Dim hBrush As Long
    Dim hOldPen As Long
    Dim hOldBrush As Long
    Dim paPoints(5) As POINTAPI
    Dim lColour As Long
    
    If bOn Then
        lColour = mlSegmentOnColour
    Else
        lColour = mlSegmentOffColour
    End If
    
    hPen = CreatePen(0, 0, lColour)
    hBrush = CreateSolidBrush(lColour)
    
    hOldPen = SelectObject(moForm.hdc, hPen)
    hOldBrush = SelectObject(moForm.hdc, hBrush)
    
    With paPoints(0)
        .x = lX
        .y = lY
    End With
    With paPoints(1)
        .x = lX + mlWidthVert \ 2
        .y = lY - mlHeightHoriz \ 2
    End With
    With paPoints(2)
        .x = lX + mlWidthHoriz - mlWidthVert \ 2
        .y = lY - mlHeightHoriz \ 2
    End With
    With paPoints(3)
        .x = lX + mlWidthHoriz
        .y = lY
    End With
    With paPoints(4)
        .x = lX + mlWidthHoriz - mlWidthVert \ 2
        .y = lY + mlHeightHoriz \ 2
    End With
    With paPoints(5)
        .x = lX + mlWidthVert \ 2
        .y = lY + mlHeightHoriz \ 2
    End With
    
    Polygon moForm.hdc, paPoints(0), UBound(paPoints) + 1
    
    Call SelectObject(moForm.hdc, hOldPen)
    Call SelectObject(moForm.hdc, hOldBrush)
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Private Sub DrawDiagonalSegment(ByVal lX As Long, ByVal lY As Long, ByVal lDirection As Long, ByVal bOn As Boolean)
    Dim hPen As Long
    Dim hBrush As Long
    Dim hOldPen As Long
    Dim hOldBrush As Long
    Dim paPoints(5) As POINTAPI
    Dim lColour As Long
    
    If bOn Then
        lColour = mlSegmentOnColour
    Else
        lColour = mlSegmentOffColour
    End If
    
    hPen = CreatePen(0, 0, lColour)
    hBrush = CreateSolidBrush(lColour)
    
    hOldPen = SelectObject(moForm.hdc, hPen)
    hOldBrush = SelectObject(moForm.hdc, hBrush)
    
    With paPoints(0)
        .x = lX
        .y = lY
    End With
    With paPoints(1)
        .x = lX + mlWidthVert \ 2
        .y = lY
    End With
    With paPoints(2)
        .x = lX + mlWidthHoriz - mlWidthVert \ 2
        .y = lY + mlHeightVert - mlHeightHoriz \ 2
    End With
    With paPoints(3)
        .x = lX + mlWidthHoriz - mlWidthVert \ 2
        .y = lY + mlHeightHoriz
    End With
    With paPoints(4)
        .x = lX + mlWidthHoriz - mlWidthVert \ 2
        .y = lY + mlHeightHoriz \ 2
    End With
    With paPoints(5)
        .x = lX + mlWidthVert \ 2
        .y = lY + mlHeightHoriz \ 2
    End With
    
    Polygon moForm.hdc, paPoints(0), UBound(paPoints) + 1
    
    Call SelectObject(moForm.hdc, hOldPen)
    Call SelectObject(moForm.hdc, hOldBrush)
    DeleteObject hPen
    DeleteObject hBrush
End Sub

