VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPad 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sTextMem As String

Private vKeyMap As Variant
Private vMatches As Variant

Private Sub Form_Load()
    vKeyMap = Array(" abcdefghijklmnopqrstuvwxyz")
    sTextMem = "its all the same to me"
End Sub

Private Sub Form_Resize()
    txtPad.Width = Me.Width
    txtPad.Height = Me.Height
End Sub

Private Function PredictString(lKey As Long, sText As String) As String
    Dim lMatch As Long
    Dim lSize As Long
    Dim lPos As Long
    Dim lMaxLen As Long
    Dim sMatch As String
    Dim bFound As Boolean
    Dim lSearchMatches As Long
    Dim sPrediction As String
    Dim sTextSequence As String
    Dim sMemSequence As String
    Dim bSorted As Boolean
    Dim vTemp As Variant
    Dim lIndex As Long
    Dim lIndex2 As Long
    
    vMatches = Array()
    
    sTextSequence = Sequence(sText)
    sMemSequence = Sequence(sTextMem)
    
    For lMatch = 1 To 27
        If Len(sText) > Len(sTextMem) Then
            lMaxLen = Len(sTextMem)
        End If
        
        For lSize = Len(sText) To 0 Step -1
            For lPos = 1 To Len(sTextMem) - lSize
                sPrediction = Mid$(sTextMem, lPos + lSize, 1)
                If InStr(vKeyMap(lKey), sPrediction) <> -1 Then
                    If Mid$(sMemSequence, lPos, lSize) = Right$(sTextSequence, lSize) Then
                        bFound = False
                        For lSearchMatches = 0 To UBound(vMatches)
                            If vMatches(lSearchMatches)(0) = Right$(sText, lSize) And vMatches(lSearchMatches)(2) = sPrediction Then
                                bFound = True
                                Exit For
                            End If
                        Next
                        If Not bFound Then
                            ReDim Preserve vMatches(UBound(vMatches) + 1) As Variant
                            vMatches(UBound(vMatches)) = Array(Mid$(sTextMem, lPos, lSize), 1, sPrediction)
                        Else
                            vMatches(lSearchMatches)(1) = vMatches(lSearchMatches)(1) + 1
                        End If
                    End If
                End If
            Next
        Next
    Next
    
    ' Sort
    While Not bSorted
        bSorted = True
        For lIndex = 0 To UBound(vMatches) - 1
            If vMatches(lIndex)(1) < vMatches(lIndex + 1)(1) Then
                vTemp = vMatches(lIndex)
                vMatches(lIndex) = vMatches(lIndex + 1)
                vMatches(lIndex + 1) = vTemp
                bSorted = False
                Exit For
            End If
        Next
        
    Wend
    
    ' Remove duplicates
    For lIndex = 0 To UBound(vMatches) - 1
        For lIndex2 = lIndex + 1 To UBound(vMatches)
            If Not IsEmpty(vMatches(lIndex)) And Not IsEmpty(vMatches(lIndex2)) Then
                If vMatches(lIndex)(2) = vMatches(lIndex2)(2) Then
                    vMatches(lIndex2) = Empty
                End If
            End If
        Next
    Next
    PredictString = vMatches(0)(0) & vMatches(0)(2)
End Function

Private Sub txtPad_KeyPress(KeyAscii As Integer)
    Dim sPrediction As String
    
    If KeyAscii >= 48 And KeyAscii <= 58 Then
        sPrediction = PredictString(KeyAscii - 48, txtPad.Text)
        KeyAscii = 0
        txtPad.Text = Left$(txtPad.Text, Len(txtPad.Text) - Len(sPrediction) + 1) & Right$(sPrediction, 1)
    End If
End Sub

Private Function Sequence(sText As String) As String
    Dim lPos As Long
    Dim lKeyIndex As Long
    Dim sChar As String
    
    For lPos = 1 To Len(sText)
        sChar = Mid$(sText, lPos, 1)
        For lKeyIndex = 0 To UBound(vKeyMap)
            If InStr(vKeyMap(lKeyIndex), sChar) <> -1 Then
                Sequence = Sequence & Chr$(lKeyIndex + 48)
            Else
                Sequence = Sequence & "x"
            End If
        Next
    Next
End Function
