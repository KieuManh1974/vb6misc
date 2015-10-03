VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim lYear As Long
    Dim lMonth As Long
    
'    For lYear = 1900 To 2000
'        Debug.Print Weekday(DateSerial(lYear, 1, 1) - lYear + 1900)
'    Next
    For lMonth = 1 To 12
        Debug.Print Weekday(DateSerial(2005, lMonth, 1) - 1, vbMonday) & " " & Format$(DateSerial(2005, lMonth, 1), "MMM")
    Next
End Sub
