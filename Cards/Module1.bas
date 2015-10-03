Attribute VB_Name = "Module1"
Option Explicit

Public Sub main()
    Dim lTrials As Long
    Dim lCount As Long
    Dim lCard1 As Long
    Dim lCard2 As Long
    
    Randomize
    
    Do
        lCard1 = Int(Rnd * 52)
        lCard2 = Int(Rnd * 52)
        
        If lCard1 <> lCard2 Then
            If lCard1 \ 4 = lCard2 \ 4 Then
                lCount = lCount + 1
            End If
            
            lTrials = lTrials + 1
        End If
        Debug.Print lCount / lTrials
    Loop
    
End Sub


