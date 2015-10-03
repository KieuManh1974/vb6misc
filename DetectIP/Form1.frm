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
   Begin VB.Timer moTimer 
      Interval        =   60000
      Left            =   1920
      Top             =   360
   End
   Begin InetCtlsObjects.Inet moInet 
      Left            =   600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      URL             =   "http://"
      RequestTimeout  =   10
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    moTimer_Timer
End Sub

Private Sub moTimer_Timer()
    Static lMinutes As Long
    Dim sResponse As String
    Dim poSendMail As vbSendMail.clsSendMail
    Dim sURL As String
    Dim lURLStart As Long
    Dim lURLEnd As Long
    Static sServerIP As String
    Dim sNewServerIP As String
    
    If lMinutes = 0 Then
        sURL = "http://www.showipaddress.com/?pp=" & Format$(Time, "HHNNSS")
        sResponse = moInet.OpenURL(sURL)
        If sResponse <> "" Then
            lURLStart = InStr(sResponse, "Remote IP Address:")
            lURLStart = InStr(lURLStart, sResponse, "<td class=""color1"">")
            lURLEnd = InStr(lURLStart, sResponse, "</td>")
            sNewServerIP = Mid$(sResponse, lURLStart + 19, lURLEnd - lURLStart - 19)
        End If
        If sServerIP <> sNewServerIP Then
            sServerIP = sNewServerIP
            Set poSendMail = New vbSendMail.clsSendMail
            poSendMail.SMTPHost = "smtp.googlemail.com"
            poSendMail.POP3Host = "pop.googlemail.com"
            poSendMail.From = "gmp@ccb.ac.uk"
            poSendMail.FromDisplayName = "GMPServer"
            poSendMail.Recipient = "guille.phillips@googlemail.com"
            poSendMail.RecipientDisplayName = "Guille Phillips"
            poSendMail.ReplyToAddress = "guille.phillips@googlemail.com"
            poSendMail.Subject = "Server IP: " & sServerIP
            'poSendMail.Attachment = "" ' 'attached file name
            poSendMail.Message = "Server IP: " & sServerIP
            
            poSendMail.UseAuthentication = True             ' Optional, default = FALSE
            poSendMail.Username = "guille.phillips@googlemail.com"                     ' Optional, default = Null String
            poSendMail.Password = "rotherhithe"
            poSendMail.SMTPPort = 465
            poSendMail.Send
            Set poSendMail = Nothing
        End If
    End If
    lMinutes = (lMinutes + 1) Mod 60
End Sub
