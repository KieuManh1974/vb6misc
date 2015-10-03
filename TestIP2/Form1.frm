VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
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
   Begin InetCtlsObjects.Inet moInet 
      Left            =   480
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sResponse As String
    Dim lValue2 As Long
    Dim lValue1 As Long
    Dim lValue0 As Long
    Dim sURL As String
    Dim fTime As Double
    
    On Error Resume Next
    
    '89.213.39.248
    For lValue2 = 0 To 255
        For lValue1 = 0 To 255
            For lValue0 = 0 To 255
                sURL = "http://89." & lValue2 & "." & lValue1 & "." & lValue0
                sResponse = ""
                moInet.Cancel
                StartCounter
                sResponse = moInet.OpenURL(sURL)
                fTime = GetCounter
                
'                If fTime > 5 Then
'                    Debug.Print sURL & "?"
'                End If
                If InStr(sResponse, "504 Gateway Time-Out") > 0 Or sResponse = "" Then
                Else
                    Debug.Print sURL
                    'Debug.Print sResponse
                End If
            Next
            DoEvents
        Next
    Next
    
End Sub

