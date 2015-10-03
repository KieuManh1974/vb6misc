Attribute VB_Name = "Module1"
Option Explicit

Private msKeys() As String
Private mlValues() As String

Private msNotEmpty As Boolean

Private Sub Main()
    Match "eeedekend?"
End Sub

Private Sub Match(sText As String)
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim lIndex3 As Long
    Dim sChar1 As String
    Dim sChar2 As String
    Dim lOffset As Long
    Dim sMatch As String
    Dim lStart As Long
    
    For lIndex1 = 2 To Len(sText)
        lOffset = -lIndex1 + 1
        lStart = 0
        sMatch = ""
        For lIndex2 = 0 To Len(sText) - lIndex1
            sChar1 = Mid$(sText, lIndex1 + lOffset + lIndex2, 1)
            sChar2 = Mid$(sText, lIndex1 + lIndex2, 1)
            
            If sChar1 <> sChar2 Then
                If sMatch <> "" Then
                    IncludeKey sMatch
                End If
                sMatch = ""
            Else
                sMatch = sMatch & sChar1
            End If
        Next
        If sMatch <> "" Then
            IncludeKey sMatch
        End If
    Next
End Sub

Private Sub IncludeKey(sKey As String)
    Dim lIndex As Long
    
    If msNotEmpty Then
        For lIndex = LBound(msKeys) To UBound(msKeys)
            If msKeys(lIndex) = sKey Then
                mlValues(lIndex) = mlValues(lIndex) + 1
                Exit Sub
            End If
        Next
        ReDim Preserve msKeys(UBound(msKeys) + 1)
        ReDim Preserve mlValues(UBound(mlValues) + 1)
        msKeys(UBound(msKeys)) = sKey
        mlValues(UBound(mlValues)) = 1
    Else
        ReDim msKeys(0)
        ReDim mlValues(0)
        
        msKeys(0) = sKey
        mlValues(0) = 1
        msNotEmpty = True
    End If

End Sub

