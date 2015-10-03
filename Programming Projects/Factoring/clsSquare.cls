VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlSquare(20, 20) As Long
Private mlFactor(1, 20) As Long
Private mlSize As Long

Public Property Get Display() As String
    Dim lColumn As Long
    Dim lRow As Long
    
    For lRow = 0 To mlSize
        For lColumn = mlSize To 0 Step -1
            Display = Display & Mid$("0x", mlSquare(lColumn, lRow) + 1, 1)
        Next
        Display = Display & vbCrLf
    Next
End Property

Public Property Let Size(lSize As Long)
    mlSize = lSize - 1
End Property

Public Property Get Size() As Long
    Size = mlSize + 1
End Property

Public Property Get Factor(lFactorIndex As Long) As Long
    Dim lIndex As Long
    
    CheckOnes
    For lIndex = 0 To mlSize
        Factor = Factor + mlFactor(lFactorIndex, lIndex) * 2 ^ lIndex
    Next
End Property

Public Sub Initialise(ByVal lNumber As Long)
    Dim lIndex As Long
    Dim vBinary As Variant
    
    vBinary = Binary(lNumber)
    For lIndex = 0 To mlSize
        mlSquare(lIndex, 0) = vBinary(lIndex)
    Next
End Sub

Private Function Binary(lNumber) As Variant
    Dim lIndex As Long
    Dim vBinary As Variant
    
    vBinary = Array()
    
    ReDim vBinary(mlSize) As Long
    
    For lIndex = 0 To mlSize
        vBinary(lIndex) = lNumber Mod 2
        lNumber = lNumber \ 2
    Next
    Binary = vBinary
End Function

Public Function Move(lDiagonal As Long) As Boolean
    Dim lTemp As Long
    Dim vOnes As Variant
    Dim lPosition As Long
    Dim bOK As Boolean
    Dim lDiagonalSize As Long
    Dim lOffset As Long
    
    lDiagonalSize = DiagonalSize(lDiagonal) - 1
    
    If lDiagonalSize = 0 Then
        Exit Function
    End If
    
    vOnes = FindOnes(lDiagonal)
    If UBound(vOnes) < lDiagonalSize And UBound(vOnes) >= 0 Then
        While Not bOK
            lPosition = vOnes(Int(Rnd * (UBound(vOnes) + 1)))
            
            If lPosition = 0 Then
                If Square(lDiagonal, 1) = 0 Then
                    bOK = True
                    Move = True
                    Square(lDiagonal, 0) = 0
                    Square(lDiagonal, 1) = 1
                End If
            ElseIf lPosition = lDiagonalSize Then
                If Square(lDiagonal, lPosition - 1) = 0 Then
                    bOK = True
                    Move = True
                    Square(lDiagonal, lPosition) = 0
                    Square(lDiagonal, lPosition - 1) = 1
                End If
            Else
                If lDiagonalSize > 1 Then
                    If Square(lDiagonal, lPosition - 1) = 0 And Square(lDiagonal, lPosition + 1) = 0 Then
                        bOK = True
                        Move = True
                        lOffset = Int(Rnd * 2) * 2 - 1
                        Square(lDiagonal, lPosition) = 0
                        Square(lDiagonal, lPosition + lOffset) = 1
                    ElseIf Square(lDiagonal, lPosition - 1) = 1 And Square(lDiagonal, lPosition + 1) = 1 Then
                    ElseIf Square(lDiagonal, lPosition - 1) = 1 Then
                        bOK = True
                        Move = True
                        Square(lDiagonal, lPosition) = 0
                        Square(lDiagonal, lPosition + 1) = 1
                    Else
                        bOK = True
                        Move = True
                        Square(lDiagonal, lPosition) = 0
                        Square(lDiagonal, lPosition - 1) = 1
                    End If
                End If
            End If
        Wend
    End If
End Function

Public Function Split(lDiagonal As Long) As Boolean
    Dim vOnes As Variant
    Dim vZeroes As Variant
    Dim lPosition1 As Long
    Dim lPosition2 As Long
    Dim lPosition3 As Long
    
    If lDiagonal < 1 Then
        Exit Function
    End If
    
    vOnes = FindOnes(lDiagonal)
    If UBound(vOnes) >= 0 Then
        vZeroes = FindZeroes(lDiagonal - 1)
        If UBound(vZeroes) >= 1 Then
            While lPosition1 = lPosition2
                lPosition1 = vZeroes(Int(Rnd * (UBound(vZeroes) + 1)))
                lPosition2 = vZeroes(Int(Rnd * (UBound(vZeroes) + 1)))
            Wend
            lPosition3 = vOnes(Int(Rnd * (UBound(vOnes) + 1)))
            Square(lDiagonal - 1, lPosition1) = 1
            Square(lDiagonal - 1, lPosition2) = 1
            Square(lDiagonal, lPosition3) = 0
            Split = True
        End If
    End If
End Function

Public Function Join(lDiagonal As Long) As Boolean
    Dim vOnes As Variant
    Dim vZeroes As Variant
    Dim lPosition1 As Long
    Dim lPosition2 As Long
    Dim lPosition3 As Long
    
    If lDiagonal >= (mlSize * 2) Then
        Exit Function
    End If
    
    vOnes = FindOnes(lDiagonal)
    If UBound(vOnes) >= 1 Then
        vZeroes = FindZeroes(lDiagonal + 1)
        If UBound(vZeroes) >= 0 Then
            While lPosition1 = lPosition2
                lPosition1 = vOnes(Int(Rnd * (UBound(vOnes) + 1)))
                lPosition2 = vOnes(Int(Rnd * (UBound(vOnes) + 1)))
            Wend
            lPosition3 = vZeroes(Int(Rnd * (UBound(vZeroes) + 1)))
            
            Square(lDiagonal, lPosition1) = 0
            Square(lDiagonal, lPosition2) = 0
            Square(lDiagonal + 1, lPosition3) = 1
            Join = True
        End If
    End If
End Function

Private Function FindZeroes(lDiagonal As Long) As Variant
    Dim lPosition As Long
    Dim vZeroes As Variant
    
    vZeroes = Array()
    For lPosition = 0 To DiagonalSize(lDiagonal) - 1
        If Square(lDiagonal, lPosition) = 0 Then
            ReDim Preserve vZeroes(UBound(vZeroes) + 1)
            vZeroes(UBound(vZeroes)) = lPosition
        End If
    Next
    FindZeroes = vZeroes
End Function

Private Function FindOnes(lDiagonal As Long) As Variant
    Dim lPosition As Long
    Dim vOnes As Variant
    
    vOnes = Array()
    For lPosition = 0 To DiagonalSize(lDiagonal) - 1
        If Square(lDiagonal, lPosition) = 1 Then
            ReDim Preserve vOnes(UBound(vOnes) + 1)
            vOnes(UBound(vOnes)) = lPosition
        End If
    Next
    FindOnes = vOnes
End Function

Private Property Get DiagonalSize(lDiagonal) As Long
    If lDiagonal <= mlSize Then
        DiagonalSize = lDiagonal + 1
    Else
        DiagonalSize = mlSize * 2 + 1 - lDiagonal
    End If
End Property

Private Property Let Square(lDiagonal As Long, lPosition As Long, lValue As Long)
    If lDiagonal <= mlSize Then
        mlSquare(lDiagonal - lPosition, lPosition) = lValue
    Else
        mlSquare(mlSize - lPosition, lDiagonal - mlSize + lPosition) = lValue
    End If
End Property

Private Property Get Square(lDiagonal As Long, lPosition As Long) As Long
    If lDiagonal <= mlSize Then
        Square = mlSquare(lDiagonal - lPosition, lPosition)
    Else
        Square = mlSquare(mlSize - lPosition, lDiagonal - mlSize + lPosition)
    End If
End Property

Public Property Let Grid(lRow As Long, lColumn As Long, lValue As Long)
    mlSquare(lRow, lColumn) = lValue
End Property

Public Property Get Grid(lRow As Long, lColumn As Long) As Long
    Grid = mlSquare(lRow, lColumn)
End Property

Private Sub CheckOnes()
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    
    For lIndex1 = 0 To mlSize
        mlFactor(0, lIndex1) = 0
        mlFactor(1, lIndex1) = 0
    Next
    
    For lIndex1 = 0 To mlSize
        For lIndex2 = 0 To mlSize
            If mlSquare(lIndex1, lIndex2) = 1 Then
                mlFactor(0, lIndex1) = mlFactor(0, lIndex1) Or 1
                mlFactor(1, lIndex2) = mlFactor(1, lIndex2) Or 1
            End If
        Next
    Next
End Sub

Public Function MismatchCount() As Long
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    
    CheckOnes
    
    For lIndex1 = 0 To mlSize
        For lIndex2 = 0 To mlSize
            If mlSquare(lIndex1, lIndex2) = 0 And mlFactor(0, lIndex1) = 1 And mlFactor(1, lIndex2) = 1 Then
                MismatchCount = MismatchCount + 1
            End If
'            If mlSquare(lIndex1, lIndex2) = 1 And (mlFactor(0, lIndex1) = 0 Or mlFactor(1, lIndex2) = 0) Then
'                MismatchCount = MismatchCount + 1
'            End If
        Next
    Next
End Function

Public Property Get Clone() As clsSquare
    Dim lRow As Long
    Dim lColumn As Long
    
    Set Clone = New clsSquare
    
    For lRow = 0 To mlSize
        For lColumn = 0 To mlSize
            Clone.Grid(lColumn, lRow) = mlSquare(lColumn, lRow)
        Next
    Next
    Clone.Size = Me.Size
End Property

Public Sub Mutate()
    Dim lOperation As Long
    Dim lDiagonal As Long
    Dim bOK As Boolean
    
    Randomize
    While Not bOK
        lDiagonal = Int(Rnd * (mlSize + 1))
        lOperation = Int(Rnd * 3)
        Select Case lOperation
            Case 0 'move
                bOK = Move(lDiagonal)
            Case 1 ' split
                bOK = Split(lDiagonal)
            Case 2 'join
                bOK = Join(lDiagonal)
        End Select
    Wend
End Sub
