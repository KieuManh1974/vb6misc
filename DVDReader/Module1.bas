Attribute VB_Name = "Module1"
Option Explicit

Private Type IFO
    descript(31) As String * 32
End Type

Sub main()
    ReadVMGIFO
End Sub

Private Sub ReadVMGIFO()
    Dim yTest As IFO
    
'    Open "D:\VIDEO_TS\VIDEO_TS.IFO" For Binary As #1
    Open "D:\VIDEO_TS\VTS_01_0.IFO" For Binary As #1
    
    Get #1, , yTest
    Close #1
    
'    Debug.Print yTest(&H104& + 3)
End Sub
