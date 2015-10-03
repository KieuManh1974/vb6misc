Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim a(65535) As Byte
   ' Dim size As Double
    
    'Form1.Show
    
    Open "d:\mpegav\avseq01.dat" For Binary As #1
    Open "c:\tv\gladiator1.avi" For Binary As #2
    
    On Error GoTo closefile
    
    While Not EOF(1)
        Get #1, , a
        Put #2, , a
        'size = size + 64
        'Form1.Label1.Caption = size & " K         "
    Wend
    
    Exit Sub
closefile:
    Close #2
    Close #1
End Sub
