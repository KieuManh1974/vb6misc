Attribute VB_Name = "Module2"
Option Explicit

Public Function ApplyKey(sCode As String) As String
    Dim lPos As Long
    
    Do
        ApplyKey = sCode
        For lPos = Len(sCode) - 1 To 1 Step -1
            Mid$(ApplyKey, lPos, 1) = CStr(Val(Mid$(ApplyKey, lPos, 1) + Val(Mid$(ApplyKey, lPos + 1, 1))) Mod 10)
        Next
    Loop Until Mid$(ApplyKey, 1, 1) <> "0"
    
End Function

Public Function ApplyInverseKey(sCode As String) As String
    Dim lPos As Long
    
    Do
        ApplyInverseKey = sCode
        For lPos = 1 To Len(sCode) - 1
            Mid$(ApplyInverseKey, lPos, 1) = CStr((Val(Mid$(ApplyInverseKey, lPos, 1)) - Val(Mid$(ApplyInverseKey, lPos + 1, 1)) + 10) Mod 10)
        Next
    Loop Until Mid$(ApplyInverseKey, 1, 1) <> "0"
End Function
