VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BackColor       =   &H00FFF4EE&
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8595
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   573
   StartUpPosition =   3  'Windows Default
   Begin MSForms.ScrollBar scrVertical 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   150
      ForeColor       =   0
      BackColor       =   12632256
      Size            =   "265;3836"
      Min             =   1
      Position        =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private mlLineSize As Single
Private id As String
Private bReady As Boolean
Private dTab As Single
Private mlCurrentFont As Long

Private msBookText As String
Private mlBookTextRows() As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SIZE
    Width As Long
    Height As Long
End Type

Private Const DT_LEFT = &H0
Private Const DT_TOP = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_NOPREFIX = &H800

Private mrectMargins As RECT
Private mbBlockResize As Boolean
    
Private mlBreakIndex As Long
Private mlLineBreakIndeces() As Long
Private mlLineBreakIndecesCounter As Long
Private mlBreaks() As Long


Private Sub Form_Activate()
'    Dim q As Single
'    For q = 0 To 255 Step 0.2
'        Me.Line (q * (Me.Width / 255), 0)-Step(0, 350), RGB(0, 0, 255 - q)
'        Me.Line (0, 350)-Step(Me.Width, 0), RGB(128, 128, 128)
'        Me.Line (0, 0)-Step(1, 0)
'        Me.ForeColor = vbWhite
'        Me.Print "Title"
'        Me.PaintPicture LoadPicture("closer.bmp"), Me.Width - 100, 350
'    Next
End Sub

Private Function GetLineSize()
    Dim sizeDimensions As SIZE
    
    GetTextExtentPoint32 Me.hDC, "Xy", 2, sizeDimensions
    mlLineSize = sizeDimensions.Height + 1
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim nSizeDiff As Single
    Dim nOriginalSize As Single

    Dim oFonts As New Collection
    Dim lFontIndex As Long
    
    oFonts.Add "Times New Roman"
    oFonts.Add "Garamond"
    oFonts.Add "Arial"
    oFonts.Add "Tahoma"
    oFonts.Add "Calibri"
        
    Select Case Chr(KeyAscii)
        Case "+"
            nSizeDiff = 0.01
            nOriginalSize = Me.Font.SIZE
            Do
                Me.Font.SIZE = Me.Font.SIZE + nSizeDiff
                nSizeDiff = nSizeDiff + 0.01
            Loop Until Me.Font.SIZE <> nOriginalSize
            GetLineSize
            ConfigureLineBreaks Me.ScaleWidth - mrectMargins.Left - mrectMargins.Right - scrVertical.Width
            Display
        Case "-"
            nSizeDiff = 0.01
            nOriginalSize = Me.Font.SIZE
            Do
                Me.Font.SIZE = Me.Font.SIZE - nSizeDiff
                nSizeDiff = nSizeDiff + 0.01
            Loop Until Me.Font.SIZE <> nOriginalSize
            GetLineSize
            ConfigureLineBreaks Me.ScaleWidth - mrectMargins.Left - mrectMargins.Right - scrVertical.Width
            Display
        Case "b", "B"
            Me.Font.Bold = Not CBool(Me.Font.Bold)
            GetLineSize
            Cls
            Display
        Case "f", "F"
            mlCurrentFont = (mlCurrentFont + 1) Mod oFonts.Count
            Me.Font.Name = oFonts(mlCurrentFont + 1)
            GetLineSize
            ConfigureLineBreaks Me.ScaleWidth - mrectMargins.Left - mrectMargins.Right - scrVertical.Width
            Display
    End Select
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim fso As New FileSystemObject
    Dim reading As TextStream
    Dim texfile As String
    Dim texinfo As File
    Dim q As Integer
    
    On Error Resume Next
    
    If Command = "" Then
        End
    End If
    id = Normal(Command)
    
    texfile = Mid(Command, 2, Len(Command) - 2)

    Set texinfo = fso.GetFile(texfile)

    Me.Caption = Left(texinfo.Name, InStr(texinfo.Name, ".") - 1)
    
    Me.Font.SIZE = GetSetting("TexReader", "Display", "FontSize", 12)
           
    GetLineSize
    dTab = TextWidth(String$(40, " "))
    
    ReDim row(0) As String

    msBookText = fso.OpenTextFile(texfile, ForReading).ReadAll
    msBookText = Replace$(msBookText, vbCrLf, vbCr)
    
    SetBreaks

    mrectMargins.Top = 50
    mrectMargins.Right = 50
    mrectMargins.Bottom = 50
    mrectMargins.Left = 50

    mbBlockResize = True
    Me.Height = GetSetting("TexReader", id, "Height", Me.Height)
    Me.Width = GetSetting("TexReader", id, "Width", Me.Width)
    Me.Top = GetSetting("TexReader", id, "Top", Me.Top)
    Me.Left = GetSetting("TexReader", id, "Left", Me.Left)
    scrVertical.Value = GetSetting("TexReader", id, "TopLine", 1)
    mbBlockResize = False
    bReady = True
    ConfigureLineBreaks Me.ScaleWidth - mrectMargins.Left - mrectMargins.Right - scrVertical.Width
    
    'mhHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0&, App.ThreadID)
End Sub

Private Sub SetScrollBar()
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lLinesPerPage As Integer
    
    lLinesPerPage = (Me.ScaleHeight - mrectMargins.Top - mrectMargins.Bottom) / mlLineSize
    
    If Button = vbLeftButton Then
        If (scrVertical.Value + lLinesPerPage) > scrVertical.Max Then
            scrVertical.Value = scrVertical.Max
        Else
            scrVertical.Value = scrVertical.Value + lLinesPerPage
        End If
    ElseIf Button = vbRightButton Then
        If (scrVertical.Value - lLinesPerPage) < scrVertical.Min Then
            scrVertical.Value = scrVertical.Min
        Else
            scrVertical.Value = scrVertical.Value - lLinesPerPage
        End If
    End If
End Sub

Private Sub Form_Paint()
    Display
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.WindowState = vbNormal
End Sub

Private Sub Form_Resize()
    scrVertical.Height = Me.ScaleHeight
    scrVertical.LargeChange = (Me.ScaleHeight - mrectMargins.Top - mrectMargins.Bottom) / mlLineSize - 1

    If Not mbBlockResize Then
        ConfigureLineBreaks Me.ScaleWidth - mrectMargins.Left - mrectMargins.Right - scrVertical.Width
        Display
    End If
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnhookWindowsHookEx(mhHook)
End Sub

Private Sub scrVertical_Change()
    If Not mbBlockResize Then
        Cls
        Display
    End If
End Sub

Private Function Normal(ByVal A As String) As String
    Dim q As Integer
    For q = 1 To Len(A)
        If Mid(A, q, 1) = "/" Or Mid(A, q, 1) = "\" Or Mid(A, q, 1) = ":" Then
            Normal = Normal & "@"
        Else
            Normal = Normal & Mid(A, q, 1)
        End If
    Next
End Function

Private Sub SetBreaks()
    Dim lTextLength As Long

    Dim lCheckIndex As Long
    Dim lCharCheck(1 To 8) As Variant
    Dim lCounter As Long
    Dim sCheckChar As String
    Dim vThisCheck As Variant
    Dim sCheckChars As String
    Dim lCheckFromPos As Long

    Dim lIndeces(1 To 8) As Long
    Dim lIndecesBounds(1 To 8) As Long
    Dim lIndexFinishedCount As Long
    Dim lNextCheckPos As Long
    Dim lNextIndexPos As Long
    Dim bFinished As Boolean

    sCheckChars = vbCr & " .,-;)]"
        
    lTextLength = Len(msBookText)

    For lCheckIndex = 1 To 8
        lCharCheck(lCheckIndex) = Array()
        sCheckChar = Mid$(sCheckChars, lCheckIndex, 1)
        lCheckFromPos = 1
        lCounter = 0
        vThisCheck = Array()
        While lCheckFromPos <= lTextLength And lCheckFromPos <> 0
            lCheckFromPos = InStr(lCheckFromPos + 1, msBookText, sCheckChar)
            If (lCheckFromPos <> 0) Then
                ReDim Preserve vThisCheck(lCounter)
                vThisCheck(lCounter) = lCheckFromPos
                lCounter = lCounter + 1
            End If
        Wend
        lCharCheck(lCheckIndex) = vThisCheck
    Next
   ' Debug.Print GetCounter


    
    ReDim mlBreaks(0)
    mlBreaks(0) = 0
    mlBreakIndex = 1
    
    ReDim mlLineBreakIndeces(0)
    mlLineBreakIndeces(0) = 0
    mlLineBreakIndecesCounter = 1
    
    For lCheckIndex = 1 To 8
        lIndecesBounds(lCheckIndex) = UBound(lCharCheck(lCheckIndex))
    Next
    
    bFinished = False
    Do
        lNextCheckPos = 0
        lNextIndexPos = -1
        lIndexFinishedCount = 0
        For lCheckIndex = 1 To 8
            If lIndecesBounds(lCheckIndex) >= lIndeces(lCheckIndex) Then
                If lNextCheckPos = 0 Then
                    lNextCheckPos = lCharCheck(lCheckIndex)(lIndeces(lCheckIndex))
                    lNextIndexPos = lCheckIndex
                Else
                    If lCharCheck(lCheckIndex)(lIndeces(lCheckIndex)) < lNextCheckPos Then
                        lNextCheckPos = lCharCheck(lCheckIndex)(lIndeces(lCheckIndex))
                        lNextIndexPos = lCheckIndex
                    End If
                End If
            Else
                lIndexFinishedCount = lIndexFinishedCount + 1
            End If
        Next
        If lNextIndexPos > -1 Then
            ReDim Preserve mlBreaks(mlBreakIndex)
            mlBreaks(mlBreakIndex) = lCharCheck(lNextIndexPos)(lIndeces(lNextIndexPos))
            If lNextIndexPos = 1 Then
                ReDim Preserve mlLineBreakIndeces(mlLineBreakIndecesCounter)
                mlLineBreakIndeces(mlLineBreakIndecesCounter) = mlBreakIndex
                mlLineBreakIndecesCounter = mlLineBreakIndecesCounter + 1
            End If
            mlBreakIndex = mlBreakIndex + 1
            lIndeces(lNextIndexPos) = lIndeces(lNextIndexPos) + 1
        End If
    Loop Until lIndexFinishedCount = 8
    
    ReDim Preserve mlBreaks(mlBreakIndex)
    mlBreaks(mlBreakIndex) = Len(msBookText) + 1
    
    ReDim Preserve mlLineBreakIndeces(mlLineBreakIndecesCounter)
    mlLineBreakIndeces(mlLineBreakIndecesCounter) = mlBreakIndex

End Sub

Private Sub ConfigureLineBreaks(lWidth As Long)
   ' lWidth = 200
    Dim lLineStartPos As Long
    Dim sizeDimensions As SIZE
    Dim lRowCount As Long
    
    Dim sTestText As String
    Dim fAverageWidth As Double
    Dim fAverageCharsInWidth As Double
    
    Dim lStartLineBreakIndex As Long
    Dim lRange As Long
    Dim lPositionRange As Long
    Dim lGuess As Long
    Dim lLineStartPosition As Long
    Dim lLineEndPosition As Long
    Dim bFullRange As Boolean
    Dim lPreviousWidth As Long
    Dim bBreakPositionFound As Boolean
    Dim lStartBreakIndex As Long
    Dim bFinished As Boolean
        
    StartCounter
    
    ReDim mlBookTextRows(0)
    mlBookTextRows(0) = 1
    
    sTestText = "The quick brown fox jumped over the lazy dog."
    GetTextExtentPoint32 Me.hDC, sTestText, Len(sTestText), sizeDimensions
    
    fAverageWidth = CDbl(sizeDimensions.Width) / CDbl(Len(sTestText))
    fAverageCharsInWidth = CDbl(lWidth) / fAverageWidth

    ReDim mlBookTextRows(0)
    mlBookTextRows(0) = 1
    lRowCount = 1
    
    Do
        lRange = mlLineBreakIndeces(lStartLineBreakIndex + 1) - lStartBreakIndex
        lPositionRange = mlBreaks(mlLineBreakIndeces(lStartLineBreakIndex + 1)) - mlBreaks(lStartBreakIndex) - 1
        
        If lPositionRange > 0 Then
            lGuess = (fAverageCharsInWidth / CDbl(lPositionRange)) * CDbl(lRange)
            If lGuess > lRange Then
                lGuess = lRange
            End If
        Else
            lGuess = 1
        End If
        
        lLineStartPosition = mlBreaks(lStartBreakIndex) + 1
        lLineEndPosition = mlBreaks(lStartBreakIndex + lGuess)
        
        GetTextExtentPoint32 Me.hDC, Mid$(msBookText, lLineStartPosition, lLineEndPosition - lLineStartPosition + 1), lLineEndPosition - lLineStartPosition + 1, sizeDimensions
        'Debug.Print "A:" & Mid$(msBookText, lLineStartPosition, lLineEndPosition - lLineStartPosition + 1)
        
        If sizeDimensions.Width = lWidth Then
           'Stop
        ElseIf sizeDimensions.Width > lWidth Then
            Do
                lGuess = lGuess - 1
                lLineEndPosition = mlBreaks(lStartBreakIndex + lGuess)
                GetTextExtentPoint32 Me.hDC, Mid$(msBookText, lLineStartPosition, lLineEndPosition - lLineStartPosition + 1), lLineEndPosition - lLineStartPosition + 1, sizeDimensions
                'Debug.Print "B:" & Mid$(msBookText, lLineStartPosition, lLineEndPosition - lLineStartPosition + 1)
            Loop Until sizeDimensions.Width <= lWidth
            If lGuess = 0 Then
                lGuess = 1
            End If
        Else
            If lGuess < lRange Then
                Do
                    lGuess = lGuess + 1
                    If lGuess <= lRange Then
                        lLineEndPosition = mlBreaks(lStartBreakIndex + lGuess)
                        GetTextExtentPoint32 Me.hDC, Mid$(msBookText, lLineStartPosition, lLineEndPosition - lLineStartPosition + 1), lLineEndPosition - lLineStartPosition + 1, sizeDimensions
                        'Debug.Print "C:" & Mid$(msBookText, lLineStartPosition, lLineEndPosition - lLineStartPosition + 1)
                    Else
                        lGuess = lGuess + 1
                        sizeDimensions.Width = lWidth + 1
                    End If
                Loop Until sizeDimensions.Width > lWidth
                lGuess = lGuess - 1
                lLineEndPosition = mlBreaks(lStartBreakIndex + lGuess)
            Else
                lGuess = lGuess + 1
            End If
        End If
        
        ReDim Preserve mlBookTextRows(lRowCount)
        mlBookTextRows(lRowCount) = lLineEndPosition + 1
        If lRowCount > 0 Then
            'Debug.Print "D:" & Mid$(msBookText, mlBookTextRows(lRowCount - 1), mlBookTextRows(lRowCount) - mlBookTextRows(lRowCount - 1))
        End If
        lRowCount = lRowCount + 1
        
        If lGuess > lRange Then
            lStartBreakIndex = lStartBreakIndex + lGuess - 1
            lStartLineBreakIndex = lStartLineBreakIndex + 1
        Else
            lStartBreakIndex = lStartBreakIndex + lGuess
        End If
    Loop Until lStartBreakIndex = mlBreakIndex
        
    scrVertical.Max = lRowCount - 1
    
    ''Debug.Print GetCounter
End Sub

Public Function Min(lValues() As Long) As Long
    Dim lValue As Long
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(lValues)
        lValue = lValues(lIndex)

        If lValue <> 0 Then
            If Min = 0 Then
                Min = lValue
            ElseIf lValue < Min Then
                Min = lValue
            End If
        End If
    Next
End Function


Private Sub Display()
    Dim lLineIndex As Long
    Dim nOffset As Single
    Dim mhBrush As Long
    Dim rectText As RECT
    
    mhBrush = CreateSolidBrush(Me.BackColor)
     
    nOffset = scrVertical.Value
    If bReady Then
        SaveSetting "TexReader", id, "TopLine", nOffset
        SaveSetting "TexReader", id, "Width", Me.Width
        SaveSetting "TexReader", id, "Height", Me.Height
        SaveSetting "TexReader", id, "Top", Me.Top
        SaveSetting "TexReader", id, "Left", Me.Left
        SaveSetting "TexReader", "Display", "FontSize", Me.Font.SIZE
    End If
    
    Cls
    SetTextColor Me.hDC, vbBlack
    For lLineIndex = nOffset - 1 To nOffset + ((Me.ScaleHeight - mrectMargins.Top - mrectMargins.Bottom) \ mlLineSize) - 1
        If lLineIndex < UBound(mlBookTextRows) Then
            rectText.Left = mrectMargins.Left + scrVertical.Width
            rectText.Right = Me.ScaleWidth - mrectMargins.Right
            rectText.Top = (lLineIndex - scrVertical.Value) * mlLineSize + mrectMargins.Top
            rectText.Bottom = rectText.Top + mlLineSize - 1
            
            If Mid$(msBookText, mlBookTextRows(lLineIndex + 1) - 1, 1) <> vbCr Then
                DrawText Me.hDC, Mid$(msBookText, mlBookTextRows(lLineIndex), mlBookTextRows(lLineIndex + 1) - mlBookTextRows(lLineIndex)), mlBookTextRows(lLineIndex + 1) - mlBookTextRows(lLineIndex), rectText, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
            Else
                If mlBookTextRows(lLineIndex + 1) > mlBookTextRows(lLineIndex) Then
                    DrawText Me.hDC, Mid$(msBookText, mlBookTextRows(lLineIndex), mlBookTextRows(lLineIndex + 1) - mlBookTextRows(lLineIndex) - 1), mlBookTextRows(lLineIndex + 1) - mlBookTextRows(lLineIndex) - 1, rectText, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
                End If
            End If
        End If
    Next lLineIndex
    
    Dim iPageNo As Integer
    Me.Font.SIZE = Me.Font.SIZE - 2
    If (((Me.ScaleHeight - mrectMargins.Top - mrectMargins.Bottom) \ mlLineSize) - 2) > 0 Then
        iPageNo = nOffset \ (((Me.ScaleHeight - mrectMargins.Top - mrectMargins.Bottom) \ mlLineSize) - 2) + 1
    Else
        iPageNo = 0
    End If
    Me.Line (Me.ScaleWidth - Me.TextWidth(iPageNo) - mrectMargins.Right, Me.ScaleHeight - mlLineSize)-Step(0, 0)
    Print iPageNo
    Me.Font.SIZE = Me.Font.SIZE + 2
End Sub
