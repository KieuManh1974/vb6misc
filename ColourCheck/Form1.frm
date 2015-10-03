VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   518
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   866
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIntensity 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "0"
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox txtSaturation 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "0"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtHue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   3840
      Width           =   615
   End
   Begin VB.VScrollBar scIntensity 
      Height          =   3495
      Left            =   600
      Max             =   255
      TabIndex        =   2
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar scSaturation 
      Height          =   3495
      Left            =   360
      Max             =   100
      TabIndex        =   1
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar scHue 
      Height          =   3495
      Left            =   120
      Max             =   360
      TabIndex        =   0
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlDropper As Long

Private Type BlockInfo
    x As Long
    y As Long
    Size As Long
    Colour As Long
End Type

Private mbiBlocks(3) As BlockInfo
Private mlSelectedBlock As Long

Private Sub Form_Activate()
    'ShowBlocks
    
    Line (100, 0)-Step(200, 200), vbRed, BF
    Line (300, 0)-Step(200, 200), RGB(128, 128, 128), BF
End Sub

Private Sub Form_Load()
    
    With mbiBlocks(0)
        .x = 80
        .y = 30
        .Size = 300
        .Colour = RGB(129, 128, 128)
    End With
    
    With mbiBlocks(1)
        .x = 400
        .y = 30
        .Size = 300
        .Colour = RGB(129, 128, 128)
    End With
    
    With mbiBlocks(2)
        .x = 80 + 100
        .y = 30 + 100
        .Size = 100
        .Colour = RGB(0, 0, 0)
    End With
    
    With mbiBlocks(3)
        .x = 400 + 100
        .y = 30 + 100
        .Size = 100
        .Colour = RGB(0, 0, 0)
    End With
    mlSelectedBlock = -1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lSelectedBlock As Long
    
    If Button = vbLeftButton Then
        lSelectedBlock = SelectBlock(x, y)
        If lSelectedBlock <> -1 Then
            mlSelectedBlock = lSelectedBlock
            mlDropper = Point(x, y)
            Block 80, 350, 50, mlDropper
            SetHSIFromRGB mbiBlocks(mlSelectedBlock).Colour And 255, mbiBlocks(mlSelectedBlock).Colour \ 256 And 255, mbiBlocks(mlSelectedBlock).Colour \ 65536 And 255
        End If
    End If
End Sub

Private Sub scHue_Change()
    txtHue.Text = scHue.Value
    SetSelectedBlockColour
End Sub

Private Sub scHue_Scroll()
    txtHue.Text = scHue.Value
    SetSelectedBlockColour
End Sub

Private Sub scIntensity_Change()
    txtIntensity.Text = scIntensity.Value
    SetSelectedBlockColour
End Sub

Private Sub scIntensity_Scroll()
    txtIntensity.Text = scIntensity.Value
    SetSelectedBlockColour
End Sub

Private Sub scSaturation_Change()
    txtSaturation.Text = scSaturation.Value
    SetSelectedBlockColour
End Sub

Private Sub scSaturation_Scroll()
    txtSaturation.Text = scSaturation.Value
    SetSelectedBlockColour
End Sub

Private Function Block(ByVal lXPos As Long, ByVal lYPos As Long, ByVal lSize As Long, ByVal lColour As Long)
    Line (lXPos, lYPos)-Step(lSize, lSize), lColour, BF
End Function

Private Sub ShowBlocks()
    Dim lBlockIndex As Long
    
    For lBlockIndex = 0 To UBound(mbiBlocks)
        Block mbiBlocks(lBlockIndex).x, mbiBlocks(lBlockIndex).y, mbiBlocks(lBlockIndex).Size, mbiBlocks(lBlockIndex).Colour
    Next
End Sub

Private Function SelectBlock(ByVal lXPos As Long, ByVal lYPos As Long) As Long
    Dim lBlockIndex As Long
    
    SelectBlock = -1
    For lBlockIndex = UBound(mbiBlocks) To 0 Step -1
        If lXPos >= mbiBlocks(lBlockIndex).x And lXPos <= (mbiBlocks(lBlockIndex).x + mbiBlocks(lBlockIndex).Size) And lYPos >= mbiBlocks(lBlockIndex).y And lYPos <= (mbiBlocks(lBlockIndex).y + mbiBlocks(lBlockIndex).Size) Then
            SelectBlock = lBlockIndex
            Exit Function
        End If
    Next
End Function

Private Sub SetSelectedBlockColour()
    If mlSelectedBlock <> -1 Then
        mbiBlocks(mlSelectedBlock).Colour = RGBFromHSI(scHue.Value, scSaturation.Value, scIntensity.Value)
        ShowBlocks
    End If
End Sub

Private Function RGBFromHSI(ByVal lHue As Long, ByVal lSaturation, ByVal lIntensity) As Long
    Dim h As Double
    Dim s As Double
    Dim i As Double
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim lCase As Long
    
    h = CDbl(lHue) * (Atn(1) * 4) / 180
    
    If h < (Atn(1) * 8 / 3) Then
        lCase = 1
    ElseIf h <= (Atn(1) * 16 / 3) Then
        lCase = 2
        h = h - Atn(1) * 8 / 3
    Else
        lCase = 3
        h = h - Atn(1) * 16 / 3
    End If
    
    s = CDbl(lSaturation) / 100
    i = CDbl(lIntensity) / 255
    
    x = i * (1 - s)
    
    y = i * (1 + (s * Cos(h) / Cos(Atn(1) * 4 / 3 - h)))
    
    z = 3 * i - (x + y)
    
    If x > 1 Then
        x = 1
    End If
    If y > 1 Then
        y = 1
    End If
    If z > 1 Then
        z = 1
    End If
    If x < 0 Then
        x = 0
    End If
    If y < 0 Then
        y = 0
    End If
    If z < 0 Then
        z = 0
    End If
    Select Case lCase
        Case 1
            RGBFromHSI = RGB(255 * y, 255 * z, 255 * x)
        Case 2
            RGBFromHSI = RGB(255 * x, 255 * y, 255 * z)
        Case 3
            RGBFromHSI = RGB(255 * z, 255 * x, 255 * y)
    End Select
End Function

Private Function SetHSIFromRGB(ByVal lRed As Long, ByVal lGreen As Long, ByVal lBlue As Long) As Long
    Dim r As Double
    Dim g As Double
    Dim b As Double
    Dim h As Double
    Dim s As Double
    Dim i As Double
    
    If (lRed + lGreen + lBlue) = 0 Then
        scHue.Value = 0
        scSaturation.Value = 0
        scIntensity.Value = 0
        Exit Function
    End If
    r = lRed / (lRed + lGreen + lBlue)
    g = lGreen / (lRed + lGreen + lBlue)
    b = lBlue / (lRed + lGreen + lBlue)
    
    If ((r - g) * (r - g) + (r - b) * (g - b)) <= 0 Then
        h = 0
    Else
        If b <= g Then
            h = ACS(0.5 * (r - g + r - b) / Sqr((r - g) * (r - g) + (r - b) * (g - b)))
        Else
            h = Atn(1) * 8 - ACS(0.5 * (r - g + r - b) / Sqr((r - g) * (r - g) + (r - b) * (g - b)))
        End If
    End If
    s = 1 - 3 * Min(r, g, b)
    i = (lRed + lGreen + lBlue) / (3 * 255)
    
    scHue.Value = h * 180 / (Atn(1) * 4)
    scSaturation.Value = s * 100
    scIntensity.Value = i * 255
    
End Function

Private Function ACS(ByVal dValue As Double) As Double
    If dValue = 1 Then
        ACS = 0
    Else
        ACS = 2 * Atn(1) - Atn(dValue / Sqr(1 - dValue * dValue))
    End If
End Function

Private Function Min(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    Dim lIndex As Long
    Dim vValues As Variant
    
    Min = 1000000#
    vValues = Array(a, b, c)
    
    For lIndex = 0 To 2
        If vValues(lIndex) < Min Then
            Min = vValues(lIndex)
        End If
    Next
    
End Function
