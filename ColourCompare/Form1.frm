VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "txtIntensity"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAbsoluteIntensity 
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtAbsoluteIntensity 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox txtIntensity 
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtIntensity 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtRGB 
      Height          =   375
      Index           =   3
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "255,255,255"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtRGB 
      Height          =   375
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "255,255,255"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtRGB 
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   3
      Text            =   "255,255,255"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtRGB 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Text            =   "255,255,255"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.VScrollBar scIntensity 
      Height          =   5415
      Index           =   1
      Left            =   7200
      Max             =   720
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.VScrollBar scIntensity 
      Height          =   5415
      Index           =   0
      Left            =   120
      Max             =   720
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private mlColours(3) As Long

Private Type RGB
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type RGBi
    Red As Double
    Green As Double
    Blue As Double
End Type

Private Sub DrawBlocks()
    Line (50, 20)-Step(300, 300), (mlColours(2)), BF
    Line (50 + 75, 20 + 75)-Step(75, 75), (mlColours(3)), BF
End Sub

Private Sub DrawBlocks2()
    Line (50, 20)-Step(150, 300), (mlColours(2)), BF
    Line (50 + 151, 20)-Step(150, 300), (mlColours(3)), BF
    'StripedSquare 50 + 151, 20, 150, 300, mlColours(3)
End Sub

Private Sub StripedSquare(ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lColour As Long)
    Dim lScan As Long
    
    For lScan = lY To lY + lHeight - 1 Step 3
        Line (lX, lScan)-Step(lWidth, 0), RGB(0, 0, 255)
        Line (lX, lScan + 1)-Step(lWidth, 0), RGB(0, 199, 0)
        Line (lX, lScan + 2)-Step(lWidth, 0), RGB(223, 0, 0)
    Next
End Sub

Private Sub Form_Activate()
    mlColours(0) = AssignColour(txtRGB(2).Text)
    mlColours(1) = AssignColour(txtRGB(3).Text)
    DrawBlocks
End Sub

Private Sub scIntensity_Change(iIndex As Integer)
    Dim rgbColour As RGB
    
    CopyMemory rgbColour, mlColours(iIndex), 4&

    mlColours(iIndex + 2) = RGB(rgbColour.Red * CLng(scIntensity(iIndex).Value) \ 360&, rgbColour.Green * CLng(scIntensity(iIndex).Value) \ 360&, rgbColour.Blue * CLng(scIntensity(iIndex).Value) \ 360&)

    CopyMemory rgbColour, mlColours(iIndex + 2), 4&
    
    'txtIntensity(iIndex).Text = Format$(0.5 * (CDbl(rgbColour.Red) / 255) ^ (1 / 0.34) + (CDbl(rgbColour.Green) / 255) ^ (1 / 0.34) + 0.328 * (CDbl(rgbColour.Blue) / 255) ^ (1 / 0.34), "0.000")
    'txtIntensity(iIndex).Text = Format$((CDbl(rgbColour.Red) / 255) ^ (1 / 0.34) + (CDbl(rgbColour.Green) / 255) ^ (1 / 0.34) + 1 * (CDbl(rgbColour.Blue) / 255) ^ (1 / 0.34), "0.000")
    
    txtAbsoluteIntensity(iIndex).Text = Format$(AbsoluteIntensity(rgbColour.Red, rgbColour.Green, rgbColour.Blue), "0.000")
    txtIntensity(iIndex).Text = Format$((scIntensity(iIndex).Value / 360) ^ (1 / 0.34), "0.000")
    txtRGB(iIndex + 2) = rgbColour.Red & "," & rgbColour.Green & "," & rgbColour.Blue
    
    DrawBlocks
End Sub

Private Function GammaIntensity(ByVal lGunIntensity As Long) As Double
    GammaIntensity = (CDbl(lGunIntensity) / 255) ^ (1 / 0.34)
End Function

Private Function AbsoluteIntensity(ByVal lRed As Long, ByVal lGreen As Long, ByVal lBlue As Long)
    AbsoluteIntensity = 0.392 * GammaIntensity(lRed) + GammaIntensity(lGreen) + 0.233 * GammaIntensity(lBlue)
End Function


Private Sub scIntensity_Scroll(iIndex As Integer)
    scIntensity_Change iIndex
End Sub

Private Sub txtRGB_Change(iIndex As Integer)
    Select Case iIndex
        Case 0, 1
            mlColours(iIndex) = AssignColour(txtRGB(iIndex).Text)
            scIntensity_Change iIndex
            DrawBlocks
    End Select
End Sub

Private Function Correct(ByVal lColour As Long) As Long
    Dim rgbColour As RGB
    Dim rgbCorrected As RGB
    
    CopyMemory rgbColour, lColour, 4&
    
    rgbCorrected.Red = 255 * (rgbColour.Red / 255) ^ (1 / 2.2)
    rgbCorrected.Green = 255 * (rgbColour.Green / 255) ^ (1 / 2.2)
    rgbCorrected.Blue = 255 * (rgbColour.Blue / 255) ^ (1 / 2.2)
    
    CopyMemory Correct, rgbCorrected, 4&
End Function

Private Function ScreenToIntensity(ByVal lValue As Double) As Double

End Function

Private Function IntensityToScreen(ByVal lValue As Double) As Double

End Function

Private Function AssignColour(ByVal sColour As String) As Long
    Dim vColours As Variant
    
    On Error Resume Next
    
    vColours = Split(sColour, ",")
    If UBound(vColours) <> 2 Then
        Exit Function
    End If
    
    AssignColour = RGB(vColours(0), vColours(1), vColours(2))
End Function

Private Function Intensity(ByVal lColour As Long) As Double
    Dim rgbColour As RGB
    
    CopyMemory rgbColour, lColour, 4&
    Intensity = 0.5 * (CDbl(rgbColour.Red) / 255) ^ (1 / 0.34) + (CDbl(rgbColour.Green) / 255) ^ (1 / 0.34) + 0.328 * (CDbl(rgbColour.Blue) / 255) ^ (1 / 0.34)
End Function

Private Function RGBi(ByVal lColour As Long) As RGBi
    Dim rgbColour As RGB
    
    CopyMemory rgbColour, lColour, 4&

    RGBi.Red = 0.5 * (CDbl(rgbColour.Red) / 255) ^ (1 / 0.34)
    RGBi.Green = (CDbl(rgbColour.Green) / 255) ^ (1 / 0.34)
    RGBi.Blue = 0.328 * (CDbl(rgbColour.Blue) / 255) ^ (1 / 0.34)
End Function

