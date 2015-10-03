Attribute VB_Name = "Module2"
Option Explicit

Public Function ApplyKey(sCode As String, lTimes As Long) As String
    Dim lPos As Long
    Dim lIndex As Long
    
    ApplyKey = sCode
    For lIndex = 1 To lTimes
        Do
            For lPos = Len(sCode) - 1 To 1 Step -1
                Mid$(ApplyKey, lPos, 1) = CStr(Val(Mid$(ApplyKey, lPos, 1) + Val(Mid$(ApplyKey, lPos + 1, 1))) Mod 10)
            Next
        Loop Until Mid$(ApplyKey, 1, 1) <> "0"
    Next
End Function

Public Function ApplyInverseKey(sCode As String, lTimes As Long) As String
    Dim lPos As Long
    Dim lIndex As Long
    
    ApplyInverseKey = sCode
    For lIndex = 1 To lTimes
        Do
            For lPos = 1 To Len(sCode) - 1
                Mid$(ApplyInverseKey, lPos, 1) = CStr((Val(Mid$(ApplyInverseKey, lPos, 1)) - Val(Mid$(ApplyInverseKey, lPos + 1, 1)) + 10) Mod 10)
            Next
        Loop Until Mid$(ApplyInverseKey, 1, 1) <> "0"
    Next
End Function
