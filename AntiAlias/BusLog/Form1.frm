VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "BusLog"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer oTimer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   120
      Top             =   240
   End
   Begin InetCtlsObjects.Inet oInet 
      Left            =   600
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moBusses As New clsBusses

Private Sub Form_Load()
    While Second(Now()) <> 0
        DoEvents
    Wend
    oTimer1.Enabled = True
    LogStops
End Sub

Private Sub Form_Terminate()
    moBusses.CleanUp True
End Sub

Private Sub oTimer1_Timer()
    LogStops
    moBusses.CleanUp
End Sub

Private Sub LogStops()
    ReadTimes "North+Street+C"
    ReadTimes "Brighton+Station+B"
    ReadTimes "Chich%2EDrive+West"
    ReadTimes "Lustrells+Vale"
    ReadTimes "Chorley+Avenue"
End Sub

Private Sub ReadTimes(sBusStop As String)
    Dim sPage As String

    Dim sSearchRoute As String
    Dim sSearchDestination As String
    Dim sSearchTime As String
    
    Dim sRoute As String
    Dim sDestination As String
    Dim sTime As String
    
    Dim lRouteStart As Long
    Dim lRouteEnd As Long
    Dim lDestinationStart As Long
    Dim lDestinationEnd As Long
    Dim lTimeStart As Long
    Dim lTimeEnd As Long
    
    On Error GoTo ReadTimesExit
    
    sSearchRoute = "<td bgcolor=""#660000"" width=""60"" align=""right"" valign=""top""><span class=""dfifahrten"">"
    sSearchDestination = "<td bgcolor=""#660000"" width=""180"" style=""WORD-BREAK:BREAK-ALL""><span class=""dfifahrten"">"
    sSearchTime = "<td bgcolor=""#660000"" width=""70"" align=""right"" valign=""top""><span class=""dfifahrten"">"
    
    sPage = oInet.OpenURL("http://buses.citytransport.org.uk/smartinfo/service/jsp/main.jsp?stName=" & sBusStop & "&olifServerId=182&autorefresh=30&default_autorefresh=30&routeId=-1&stopId=" & sBusStop & "&allLines=y&optTime=now&time=&nRows=10")
    
    lRouteStart = InStr(sPage, sSearchRoute)
    While lRouteStart <> 0
        lRouteEnd = InStr(lRouteStart, sPage, "</span>")
        lDestinationStart = InStr(lRouteEnd, sPage, sSearchDestination)
        lDestinationEnd = InStr(lDestinationStart, sPage, "</span")
        lTimeStart = InStr(lDestinationEnd, sPage, sSearchTime)
        lTimeEnd = InStr(lTimeStart, sPage, "</span>")
        sRoute = Mid$(sPage, lRouteStart + Len(sSearchRoute), lRouteEnd - (lRouteStart + Len(sSearchRoute)))
        sDestination = Mid$(sPage, lDestinationStart + Len(sSearchDestination), lDestinationEnd - (lDestinationStart + Len(sSearchDestination)))
        sTime = Mid$(sPage, lTimeStart + Len(sSearchTime), lTimeEnd - (lTimeStart + Len(sSearchTime)))
        
        If Not IsDate(sTime) Then
            moBusses.AddBus sRoute, sBusStop, sDestination, ClearSeconds(Val(sTime) / 1440 + Now()), Val(sTime) / 1440, False
        Else
            moBusses.AddBus sRoute, sBusStop, sDestination, Expectedtime(sTime), Expectedtime(sTime) - Now, True
        End If
        lRouteStart = InStr(lRouteStart + 1, sPage, sSearchRoute)
    Wend
    
ReadTimesExit:
End Sub

Private Function Expectedtime(sTime As String) As Date
    Dim sTimeNow As String
    Dim sNewTime As String
    Dim dNow As Date
    Dim lTimeNowHour As Long
    Dim lTimeNowMin As Long
    Dim lTimeHour As Long
    Dim lTimeMin As Long
    
    dNow = Now()
    
    lTimeNowHour = Format$(dNow, "HH")
    lTimeNowMin = Format$(dNow, "NN")
    lTimeHour = Format$(CDate(sTime), "HH")
    lTimeMin = Format$(CDate(sTime), "NN")
    
    If (lTimeHour * 60 + lTimeMin) < (lTimeNowHour * 60 + lTimeNowMin - 60) Then
        Expectedtime = (((lTimeHour + 24) * 60 + lTimeMin) - (lTimeNowHour * 60 + lTimeNowMin)) / 1440 + dNow
    Else
        Expectedtime = (((lTimeHour) * 60 + lTimeMin) - (lTimeNowHour * 60 + lTimeNowMin)) / 1440 + dNow
    End If
    Expectedtime = ClearSeconds(Expectedtime)
End Function

Private Function ClearSeconds(dTime As Date) As Date
    ClearSeconds = DateSerial(Year(dTime), Month(dTime), Day(dTime)) + TimeSerial(Hour(dTime), Minute(dTime), 0)
End Function
