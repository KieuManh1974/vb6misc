Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    Dim iYear As Double
    Dim dblEaster As Double
    Dim dblEasterNew As Double
    Dim dblDiff As Double
    
    dblEaster = CDbl(EasterDay(1899))
    
    For iYear = 1900 To 2100
        dblEasterNew = CDbl(EasterDay(iYear))
        If EasterDay(iYear) <> EasterDay2(iYear) Then
            Debug.Print iYear
        End If
        Debug.Print ConvertToBinary(EasterDay2(iYear) - CDate("23/3/" & iYear), 6)
        
'        dblDiff = dblEasterNew - dblEaster
'        If dblDiff = 357 Then
'            Debug.Print iYear - 1 & " " & dblDiff & "  "
'        End If
'        dblEaster = dblEasterNew
    Next
End Sub

' This function will provide the date of Easter for the given year
Public Function EasterDay(ByVal dblYear As Double) As Date
    Dim C As Long
    Dim N As Long
    Dim K As Long
    Dim I As Long
    Dim J As Long
    Dim L As Long
    Dim M As Long
    Dim D As Long
    Dim Y As Long
    
    Y = dblYear
    
    C = Int(Y / 100)
    N = Y - 19 * Int(Y / 19)
    K = Int((C - 17) / 25)
    I = C - Int(C / 4) - Int((C - K) / 3) + 19 * N + 15
    I = I - 30 * Int(I / 30)
    I = I - Int(I / 28) * (1 - Int(I / 28) * Int(29 / (I + 1)) * Int((21 - N) / 11))
    J = Y + Int(Y / 4) + I + 2 - C + Int(C / 4)
    J = J - 7 * Int(J / 7)
    L = I - J
    M = 3 + Int((L + 40) / 44)
    D = L + 28 - 31 * Int(M / 4)
    
    EasterDay = CDate(D & "/" & M & "/" & Y)
    
End Function


' This function will provide the date of Easter for the given year
Public Function EasterDay2(ByVal dblYear As Double) As Date
    Dim C As Long
    Dim N As Long
    Dim K As Long
    Dim I As Long
    Dim J As Long
    Dim L As Long
    Dim M As Long
    Dim D As Long
    Dim Y As Long
    
    Y = dblYear
    
    N = Y Mod 19
    K = Int((Int(Y / 100) - 17) / 25)
    I = Int(Y / 100) - Int(Y / 400) - Int((Int(Y / 100) - K) / 3) + 19 * N + 15
    I = I Mod 30
    I = I - Int(I / 28) * (1 - Int(I / 28) * Int(29 / (I + 1)) * Int((21 - N) / 11))
    J = Y + Int(Y / 4) + I + 2 - Int(Y / 100) + Int(Y / 400)
    J = J Mod 7
    L = I - J
    M = 3 + Int((L + 40) / 44)
    D = L + 28 - 31 * Int(M / 4)
    
    EasterDay2 = CDate(D & "/" & M & "/" & Y)
    
End Function

Public Function ConvertToBinary(ByVal iNumber As Integer, ByVal iLength As Integer) As String
    Dim iIndex As Long
    
    For iIndex = 1 To iLength
        ConvertToBinary = (iNumber Mod 2) & ConvertToBinary
        iNumber = Int(iNumber / 2)
    Next
End Function
