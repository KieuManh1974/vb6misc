VERSION 5.00
Begin VB.Form frmTester 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mini Tester"
   ClientHeight    =   1125
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAnswer 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblSum2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblSum1 
      Alignment       =   1  'Right Justify
      Caption         =   "SEVEN SEVEN SEVEN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "frmTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvSum As Variant

Private Sub ShowSum()
    Randomize
    mvSum = CreateSum
    
    Select Case glConfiguration
        Case 0
            lblSum2.Left = lblSum1.Left
            lblSum2.Top = 480
            txtAnswer.Left = lblSum1.Left
            txtAnswer.Top = 840
            Me.Width = 3810
            Me.Height = 2745
            cmdNext.Top = 1440
            cmdNext.Left = 2400
        Case 1
            lblSum2.Top = lblSum1.Top
            lblSum2.Left = 3720
            txtAnswer.Top = lblSum1.Top
            txtAnswer.Left = 7200
            Me.Width = 10665
            Me.Height = 1785
            cmdNext.Top = 600
            cmdNext.Left = 9240
    End Select
    
    Select Case glSpell
        Case 0
            lblSum1 = mvSum(0)
            lblSum2 = mvSum(1)
        Case 1, 2
            lblSum1 = Spell(mvSum(0))
            lblSum2 = Spell(mvSum(1))
        Case 2
    End Select
End Sub

Private Function CreateSum()
    Dim lIndex As Long
    Dim lDig1 As Long
    Dim lDig2 As Long
    Dim sNumber1 As String
    Dim sNumber2 As String
    Dim lColumnCarry As Long
    Dim lRepeat As Long
    
    lColumnCarry = Abs(glCarryPropagation = 1)
    For lIndex = 1 To glNumDigits
        lRepeat = 0
again:
        lDig1 = Int(Rnd * 10)
        lDig2 = Int(Rnd * 10)
        lRepeat = lRepeat + 1
        
        If lRepeat > 1000 Then
            CreateSum = Array("Err", "Err")
            Exit Function
        End If
        Select Case glParity
            Case -1 'mixed
            Case 0 ' both even
                If (lDig1 Mod 2) = 1 Or (lDig2 Mod 2) = 1 Then
                    GoTo again
                End If
            Case 1 'both odd
                If (lDig1 Mod 2) = 0 Or (lDig2 Mod 2) = 0 Then
                    GoTo again
                End If
            Case 2 'odd / even
                If ((lDig1 + lDig2) Mod 2) = 0 Then
                    GoTo again
                End If
        End Select
        
        If glDistance <> -1 Then
            If Distance(lDig1, lDig2) <> glDistance Then
                GoTo again
            End If
        End If
        
        Select Case glCarry
            Case -1
                Select Case Mid$(glCarryPattern, glNumDigits - lIndex + 1, 1)
                    Case "0"
                        If (lDig1 + lDig2 + lColumnCarry) > 9 Then
                            GoTo again
                        End If
                    Case "1"
                        If (lDig1 + lDig2 + lColumnCarry) < 10 Then
                            GoTo again
                        End If
                    Case Else
                End Select
            Case 0
                If (lDig1 + lDig2 + lColumnCarry) = 10 Then
                    GoTo again
                End If
            Case 1 ' No
                If (lDig1 + lDig2 + lColumnCarry) > 9 Then
                    GoTo again
                End If
            Case 2 ' Yes
                If (lDig1 + lDig2 + lColumnCarry) < 10 Then
                    GoTo again
                End If
        End Select
        
        Select Case glCentres
            Case -1
            Case 0 ' No
                Select Case lColumnCarry
                    Case 0
                        If (lDig1 Mod 5) = 0 Or (lDig2 Mod 5) = 0 Then
                            GoTo again
                        End If
                    Case 1
                        If (lDig1 Mod 5) = 4 Or (lDig2 Mod 5) = 4 Then
                            GoTo again
                        End If
                End Select
            Case 1 ' Yes: At least one of the digits is a centre
                Select Case lColumnCarry
                    Case 0
                        If (lDig1 Mod 5) <> 0 And (lDig2 Mod 5) <> 0 Then
                            GoTo again
                        End If
                    Case 1
                        If (lDig1 Mod 5) <> 4 And (lDig2 Mod 5) <> 4 Then
                            GoTo again
                        End If
                End Select
        End Select
                
        sNumber1 = lDig1 & sNumber1
        sNumber2 = lDig2 & sNumber2
        Select Case glCarryPropagation
            Case -1
                lColumnCarry = (lDig1 + lDig2 + lColumnCarry) \ 10
            Case 0
                lColumnCarry = 0
            Case 1
                lColumnCarry = 1
        End Select
    Next
    
    CreateSum = Array(sNumber1, sNumber2)
End Function

Private Function Position(ByVal lDig As Long) As Long
    Dim sNum As String
    sNum = "2486"
    
    If lDig Mod 2 = 1 Then
        lDig = (lDig + 5) Mod 10
    End If
    
    Position = InStr(sNum, CStr(lDig)) - 1
End Function

Private Function Distance(ByVal lDig1 As Long, ByVal lDig2 As Long) As Long
    Distance = Abs(Position(lDig1) - Position(lDig2))
    If Distance = 3 Then
        Distance = 1
    End If
End Function

Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
    If txtAnswer.ForeColor = vbRed Then
        txtAnswer.Text = ""
    End If
    txtAnswer.ForeColor = vbBlack
    If KeyAscii = 13 Then
        cmdNext_Click
    End If
End Sub

Private Sub cmdNext_Click()
    Dim bNext As Boolean
    
    If Not bNext Then
        If AddResult(mvSum(0), mvSum(1), glCarryPropagation, glCarryOverflow) <> Val(txtAnswer.Text) Then
            txtAnswer.ForeColor = vbRed
            txtAnswer.SetFocus
        Else
            bNext = True
        End If
    End If
    If bNext Then
        txtAnswer.ForeColor = vbBlack
        ShowSum
        txtAnswer.Text = ""
        txtAnswer.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    ShowSum
    txtAnswer.SetFocus
End Sub

Private Sub Form_Load()
    glNumDigits = GetSetting("AddTester", "Options", "NumDigits", 3)
    glParity = GetSetting("AddTester", "Options", "Parity", 3)
    glCarry = GetSetting("AddTester", "Options", "Carry", 0)
    glCarryPropagation = GetSetting("AddTester", "Options", "CarryPropagation", 0)
    glCarryOverflow = GetSetting("AddTester", "Options", "CarryOverflow", 0)
    glCarryPattern = GetSetting("AddTester", "Options", "CarryPattern", "XXX")
    glDistance = GetSetting("AddTester", "Options", "Distance", -1)
    glCentres = GetSetting("AddTester", "Options", "Centre", 1)
    glConfiguration = GetSetting("AddTester", "Options", "Configuration", 0)
    glSpell = GetSetting("AddTester", "Options", "Spell", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "AddTester", "Options", "NumDigits", glNumDigits
    SaveSetting "AddTester", "Options", "Parity", glParity
    SaveSetting "AddTester", "Options", "Carry", glCarry
    SaveSetting "AddTester", "Options", "CarryPropagation", glCarryPropagation
    SaveSetting "AddTester", "Options", "CarryOverflow", glCarryOverflow
    SaveSetting "AddTester", "Options", "CarryPattern", glCarryPattern
    SaveSetting "AddTester", "Options", "Distance", glDistance
    SaveSetting "AddTester", "Options", "Centre", glCentres
    SaveSetting "AddTester", "Options", "Configuration", glConfiguration
    SaveSetting "AddTester", "Options", "Spell", glSpell
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub


Public Function AddResult(ByVal sNum1 As String, ByVal sNum2 As String, ByVal lPropagateCarry As Long, ByVal lOverflow As Long) As String
    Dim lIndex As Long
    Dim lCarry As Long
    Dim lDig1 As Long
    Dim lDig2 As Long
    
    lCarry = Abs(lPropagateCarry = 1)
    
    For lIndex = Len(sNum1) To 1 Step -1
        lDig1 = Val(Mid$(sNum1, lIndex, 1))
        lDig2 = Val(Mid$(sNum2, lIndex, 1))
        AddResult = (lDig1 + lDig2 + lCarry) Mod 10 & AddResult
        Select Case lPropagateCarry
            Case -1
                lCarry = (lDig1 + lDig2 + lCarry) \ 10
            Case 0
                lCarry = 0
            Case 1
                lCarry = 1
        End Select
    Next
    If lOverflow = 0 And lCarry = 1 Then
        AddResult = "1" & AddResult
    End If
End Function

Private Function Spell(ByVal sNumber As String) As String
    Dim vNumbers As Variant
    Dim lIndex As Long
    
    vNumbers = Array("zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine")
    For lIndex = 1 To Len(sNumber)
        Spell = Spell & " " & vNumbers(Val(Mid$(sNumber, lIndex, 1)))
    Next
    Spell = Mid$(Spell, 2)
End Function
