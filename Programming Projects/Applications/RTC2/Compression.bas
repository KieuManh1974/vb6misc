Attribute VB_Name = "Compression"
Option Explicit

Public Function Compress(sText As String) As String
    Dim lSize As Long
    Dim vSub As Variant
    Dim lOffset As Long
    Dim lStep As Long
    Dim oSubs As New Dictionary
    Dim sKey As String
    Dim vCount As Variant
    Dim oCounts As New Dictionary
    Dim vMetaCount As Variant
    Dim vKey As Variant
    Dim sCommonKey As String
    Dim lMaxChars As Long
    Dim lMaxSize As Long
    Dim lChars As Long
    
    For lSize = Len(sText) \ 2 To 1 Step -1
        For lStep = 1 To Len(sText) - lSize + 1
            sKey = Mid$(sText, lStep, lSize)
            If Not oSubs.Exists(sKey) Then
                vCount = 1
                oSubs.Add sKey, vCount
            Else
                vCount = oSubs.Item(sKey)
                oSubs.Item(sKey) = vCount + 1
            End If
            'Debug.Print Mid$(sText, lStep, lSize)
        Next
    Next
    
    ' Find most common substring count
    lMaxChars = 0
    lMaxSize = 0
    For Each vKey In oSubs.Keys
        lChars = oSubs.Item(CStr(vKey))
        
        If lChars > lMaxChars Then
            lMaxChars = lChars
            lMaxSize = Len(vKey)
            sCommonKey = vKey
        ElseIf lChars = lMaxChars And Len(vKey) > lMaxSize Then
            lMaxSize = Len(vKey)
            sCommonKey = vKey
        End If
    Next
    
    Dim vSet As Variant
    Dim lPos As Long
    Dim bStart As Boolean
    Dim bEnd As Boolean
    Dim sSet As String
    
    vSet = Array()
    
    ' Build set
    For lPos = 1 To Len(sText)
        If Mid$(sText, lPos, lMaxSize) = sCommonKey Then
            If lPos = 1 Then
                bStart = True
            End If
            If lPos = (Len(sText) - Len(sCommonKey) + 1) Then
                bEnd = True
            End If
            lPos = lPos + lMaxSize - 1
            ReDim Preserve vSet(UBound(vSet) + 1)
            vSet(UBound(vSet)) = sSet
            sSet = ""
        Else
            sSet = sSet & ReplaceChar(Mid$(sText, lPos, 1))
        End If
    Next
    sSet = Join(vSet, ",")
    If Not bStart Then
    
        If bEnd Then
            Compress = "{" & sSet & "}" & ReplaceChar(sCommonKey)
        Else
            'Compress = ReplaceChar(sCommonKey) & "{" & sSet & "}"
        End If
    Else
    
        If Not bEnd Then
            Compress = ReplaceChar(sCommonKey) & "{" & sSet & "}"
        Else
            'Compress = "{" & sSet & "}" & ReplaceChar(sCommonKey)
        End If
    End If
End Function

Private Function ReplaceChar(sChars As String) As String
    ReplaceChar = Replace$(sChars, "|", "||")
    ReplaceChar = Replace$(ReplaceChar, ",", "|,")
    ReplaceChar = Replace$(ReplaceChar, "$", "|$")
    ReplaceChar = Replace$(ReplaceChar, ":", "|:")
    ReplaceChar = Replace$(ReplaceChar, "{", "|{")
    ReplaceChar = Replace$(ReplaceChar, "}", "|}")
    ReplaceChar = Replace$(ReplaceChar, "@", "|@")
    ReplaceChar = Replace$(ReplaceChar, "~", "|~")
    ReplaceChar = Replace$(ReplaceChar, "#", "|#")
    ReplaceChar = Replace$(ReplaceChar, "^", "|^")
    ReplaceChar = Replace$(ReplaceChar, vbCrLf, "^")
    ReplaceChar = Replace$(ReplaceChar, vbCr, "^")
    ReplaceChar = Replace$(ReplaceChar, vbLf, "^")
    ReplaceChar = Replace$(ReplaceChar, vbTab, "%")
End Function
