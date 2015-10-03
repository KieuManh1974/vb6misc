VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1200
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "10.10.2.44"
      RemotePort      =   81
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'10.10.2.44

Option Explicit

Private oConMySQL As Object
Private lIntervalMinutes As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function UserName() As String
    Dim strBufferString As String
    Dim lngResult As Long
    Const MAX_BUFFER_LENGTH = 100
    
    strBufferString = String(MAX_BUFFER_LENGTH, " ")
    lngResult = GetUserName(strBufferString, MAX_BUFFER_LENGTH)
    UserName = Split(strBufferString, Chr$(0))(0)
End Function

Private Sub Form_Load()
    Dim vCommand As Variant
    
    If Command$ <> "" Then
        vCommand = Split(Command$, " ")
        Winsock1.RemoteHost = vCommand(0)
        lIntervalMinutes = Val(vCommand(1))
    End If

    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Static lCount As Long

    On Error Resume Next
    If lCount = 0 Then
        If lIntervalMinutes > 0 Then
            Timer1.Interval = 60000
        Else
            Timer1.Interval = 10000
        End If
            Winsock1.SendData Winsock1.LocalHostName & "|" & Winsock1.LocalIP & "|" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "|" & UserName & "|" & lIntervalMinutes
        lCount = 1
        Exit Sub
    End If

    lCount = lCount + 1
    If (lCount - 1) >= lIntervalMinutes Then
        lCount = 1
'         Send ident packet
        Winsock1.SendData Winsock1.LocalHostName & "|" & Winsock1.LocalIP & "|" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "|" & UserName & "|" & lIntervalMinutes
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim vComputerInfo As Variant
    Dim sMachineID As String
    Dim sMachineIP As String
    Dim sTimeStamp As String
    Dim sTime As String
    Dim sDate As String
    Dim sUsername As String
    Dim lInterval As Long
    Dim oMachineUsage As Object
    Dim oSession As Object
    Dim oUpdate As Object
    Dim fInterval As Double
    Dim sSQL As String
    Dim bAddNew As Boolean
    Dim lSessionID As Long
    Dim lOnOffIndex As Long
    
    On Error GoTo ExitSub
    
    Winsock1.GetData vComputerInfo, vbString
    vComputerInfo = Split(vComputerInfo, "|")
    sMachineID = vComputerInfo(0)
    sMachineIP = vComputerInfo(1)
    sTimeStamp = vComputerInfo(2)
    sUsername = vComputerInfo(3)
    lInterval = vComputerInfo(4)

    ' Server is sending this, so disable Ident packets
    If sMachineIP = Winsock1.LocalIP Then
        Timer1.Enabled = False
        Exit Sub
    End If
    
    Debug.Print "Packet: " & sMachineID & " " & sMachineIP & " " & sTimeStamp & " " & sUsername & " " & lInterval
    
    Set oConMySQL = CreateObject("ADODB.Connection")
    oConMySQL.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=intranet.ccb.ac.uk;DATABASE=rmlognew;USER=gmp;PASSWORD=rotherhithe;"
    
    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     machine_on.SessionID,"
    sSQL = sSQL & "     machine_off.TimeStamp"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     machine_on"
    sSQL = sSQL & "     INNER JOIN"
    sSQL = sSQL & "     machine_off"
    sSQL = sSQL & "     ON machine_on.SessionID = machine_off.SessionID"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & "     MachineID = '" & sMachineID & "'"
    sSQL = sSQL & "     AND ServerIP = '" & Winsock1.LocalIP & "'"
    sSQL = sSQL & "     AND UserName = '" & sUsername & "'"
    sSQL = sSQL & " ORDER BY"
    sSQL = sSQL & "     TimeStamp DESC"
    sSQL = sSQL & " LIMIT 1"
    Set oMachineUsage = CreateObject("ADODB.Recordset")
    
    oMachineUsage.Open sSQL, oConMySQL, 0, 3

    Debug.Print "Find"
    
    If oMachineUsage.EOF Then
        bAddNew = True
        Debug.Print "No Entry"
    Else
        fInterval = IIf(CDbl(lInterval) = 0, 10, 60 * CDbl(lInterval)) / 86400
        If Not IsNull(oMachineUsage!TimeStamp) Then
            If CDate(sTimeStamp) > (CDate(oMachineUsage!TimeStamp) + fInterval * 3) Then
                bAddNew = True
                Debug.Print "Seesion Timeout"
            End If
        End If
    End If
    
    If bAddNew Then
        sSQL = ""
        sSQL = sSQL & " SELECT"
        sSQL = sSQL & "     SessionID+1 NewSessionID"
        sSQL = sSQL & " FROM"
        sSQL = sSQL & "     SessionCounter"
        Set oSession = CreateObject("ADODB.Recordset")
        oSession.CursorLocation = 3 ' adUseClient
        oSession.Open sSQL, oConMySQL, 0, 3
        lSessionID = IIf(IsNull(oSession.Fields("NewSessionID")), 1, oSession.Fields("NewSessionID"))
        oSession.Close
        
        sSQL = ""
        sSQL = sSQL & " UPDATE"
        sSQL = sSQL & "     SessionCounter"
        sSQL = sSQL & " SET"
        sSQL = sSQL & "     SessionID = " & lSessionID
        oSession.CursorLocation = 3 ' adUseClient
        oSession.Open sSQL, oConMySQL, 0, 3
        
        Set oUpdate = CreateObject("ADODB.Recordset")
        
        sSQL = ""
        sSQL = sSQL & " INSERT INTO"
        sSQL = sSQL & "     machine_on"
        sSQL = sSQL & "     (MachineID,"
        sSQL = sSQL & "     MachineIP,"
        sSQL = sSQL & "     TimeStamp,"
        sSQL = sSQL & "     ServerIP,"
        sSQL = sSQL & "     UserName,"
        sSQL = sSQL & "     SessionID)"
        sSQL = sSQL & " VALUES"
        sSQL = sSQL & "     ('" & sMachineID & "',"
        sSQL = sSQL & "     '" & sMachineIP & "',"
        sSQL = sSQL & "     '" & sTimeStamp & "',"
        sSQL = sSQL & "     '" & Winsock1.LocalIP & "',"
        sSQL = sSQL & "     '" & sUsername & "',"
        sSQL = sSQL & "     '" & lSessionID & "'"
        sSQL = sSQL & "     )"
        
        oUpdate.Open sSQL, oConMySQL, 0, 3
        
        sSQL = ""
        sSQL = sSQL & " INSERT INTO"
        sSQL = sSQL & "     machine_off"
        sSQL = sSQL & "     (TimeStamp,"
        sSQL = sSQL & "     SessionID)"
        sSQL = sSQL & " VALUES"
        sSQL = sSQL & "     ('" & sTimeStamp & "',"
        sSQL = sSQL & "     '" & lSessionID & "'"
        sSQL = sSQL & "     )"
        
        oUpdate.Open sSQL, oConMySQL, 0, 3
        Debug.Print "Create " & sTimeStamp & " " & lSessionID
    Else
        Set oUpdate = CreateObject("ADODB.Recordset")
        sSQL = ""
        sSQL = sSQL & " UPDATE"
        sSQL = sSQL & "     machine_off"
        sSQL = sSQL & " SET"
        sSQL = sSQL & "     TimeStamp = '" & sTimeStamp & "' "
        sSQL = sSQL & " WHERE"
        sSQL = sSQL & "     SessionID = " & oMachineUsage!SessionID
        
        oUpdate.Open sSQL, oConMySQL, 0, 3
        Debug.Print "Update " & sTimeStamp & " " & oMachineUsage!SessionID
    End If
    
    oMachineUsage.Close
    oConMySQL.Close
    
ExitSub:
    Set oSession = Nothing
    Set oMachineUsage = Nothing
    Set oConMySQL = Nothing
End Sub

