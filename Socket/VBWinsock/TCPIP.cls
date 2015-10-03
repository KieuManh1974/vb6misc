VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mwsaData As WSA_DATA
Private mlngSocket As Long

Private mlngRemotePort As Long

Private mstrLocalHostIP As String
Private msaLocalAddr As SOCK_ADDR

Private mstrRemoteHostIP As String
Private msaRemoteAddr As SOCK_ADDR

Private mlngLastErrorCode As Long
Private mstrLastErrorDesc As String

Public Function OpenConnection() As Boolean
    Dim WSAResult As Long
    
    OpenConnection = False
     
    'Initialize Winsock API
    WSAResult = WSAStartup(&H101, mwsaData)
    If WSAResult <> WSANOERROR Then
       SetLastErrorCode "Error  in OpenConnection::WSAStartup"
       Exit Function
    End If
    
     'Create new socket
     mlngSocket = socket(PF_INET, SOCK_STREAM, 0)
     If (mlngSocket = INVALID_SOCKET) Then
        SetLastErrorCode "Error in OpenConnection::socket"
        Exit Function
     End If
    
     'Bind socket to LocalHostIP
     msaLocalAddr.sin_family = PF_INET
     msaLocalAddr.sin_port = 0
     msaLocalAddr.sin_addr.S_addr = inet_addr(mstrLocalHostIP)
     If (msaLocalAddr.sin_addr.S_addr = INADDR_NONE) Then
        SetLastErrorCode "Error in OpenConnection::inet_addr"
        Exit Function
     End If
     WSAResult = bind(mlngSocket, msaLocalAddr, Len(msaLocalAddr))
     If (WSAResult = SOCKET_ERROR) Then
        SetLastErrorCode "Error in OpenConnection::bind"
        Exit Function
     End If
    
     'Connect with remote host
     msaRemoteAddr.sin_family = PF_INET
     msaRemoteAddr.sin_port = htons(mlngRemotePort)
     msaRemoteAddr.sin_addr.S_addr = inet_addr(mstrRemoteHostIP)
     If (msaLocalAddr.sin_addr.S_addr = INADDR_NONE) Then
        SetLastErrorCode "Error in OpenConnection::inet_addr=INADDR_NONE"
        Exit Function
     End If
     msaRemoteAddr.sin_zero(0) = 0
     WSAResult = connect(mlngSocket, msaRemoteAddr, Len(msaRemoteAddr))
     If (WSAResult = SOCKET_ERROR) Then
        SetLastErrorCode "Error in OpenConnection::connect"
     Else
        OpenConnection = True
     End If
End Function

Public Function ReceiveData(strData, lngBytesReceived) As Boolean
    Const MAX_BUFF_SIZE = 10000
    Dim Buff(0 To MAX_BUFF_SIZE) As Byte
    Dim WSAResult As Long
    
    WSAResult = recv(mlngSocket, Buff(0), MAX_BUFF_SIZE, 0)
    If WSAResult = SOCKET_ERROR Then
        SetLastErrorCode "Error in RecvData::recv"
        strData = ""
        lngBytesReceived = 0
        ReceiveData = False
    Else
        lngBytesReceived = WSAResult
        Buff(lngBytesReceived) = 0
        strData = Left(StrConv(Buff(), vbUnicode), lngBytesReceived)
        ReceiveData = True
    End If
End Function

Public Function SendData(ByVal strData As String) As Boolean
    Dim WSAResult As Long, i As Long, l As Long
      
    l = Len(strData)
    ReDim Buff(l + 1) As Byte
    
    For i = 1 To l
        Buff(i - 1) = Asc(Mid(strData, i, 1))
    Next
    Buff(l) = 0

    WSAResult = send(mlngSocket, Buff(0), l, 0)
    If WSAResult = SOCKET_ERROR Then
        SetLastErrorCode "Error en SendData::send"
        SendData = False
    Else
        SendData = True
    End If
End Function

Public Function IsDataAvailable() As Boolean
    Dim readfds As FD_SET, writefds As FD_SET, exceptfds As FD_SET
    Dim timeout As TIME_VAL
    Dim lngResult As Long, nfds As Long
    
    nfds = 0
    timeout.tv_sec = 1
    timeout.tv_usec = 0
    
    readfds.fd_count = 1
    readfds.fd_array(0) = mlngSocket
    writefds.fd_count = 0
    exceptfds.fd_count = 0
    
    lngResult = sselect(nfds, readfds, writefds, exceptfds, timeout)
    If lngResult = SOCKET_ERROR Then
        SetLastErrorCode "Error in IsDataAvailable::select"
        IsDataAvailable = False
    Else
        If lngResult > 0 Then IsDataAvailable = True Else IsDataAvailable = False
    End If
End Function

Public Sub ShutdownConnection()
     Dim lngResult As Long
     lngResult = shutdown(mlngSocket, 2)
     lngResult = closesocket(mlngSocket)
     lngResult = WSACancelBlockingCall
     lngResult = WSACleanup
     mlngSocket = 0
End Sub

Private Sub SetLastErrorCode(ByVal strAdditionalInfo As String)
    
    mlngLastErrorCode = WSAGetLastError
    Select Case mlngLastErrorCode
    Case INADDR_NONE
                            mstrLastErrorDesc = "INNADDR_NONE: The passed-in string does not contain a legitimate Internet address (for example, if a portion of an 'a.b.c.d' address exceeds 255), inet_addr returns the value INADDR_NONE"
    Case WSASYSNOTREADY
                            mstrLastErrorDesc = "WSASYSNOTREADY: Indicates that the underlying network subsystem is not ready for network communication"
    Case WSAVERNOTSUPPORTED
                            mstrLastErrorDesc = "WSAVERNOTSUPPORTED: The version of Windows Sockets support requested is not provided by this particular Windows Sockets implementation"
    Case WSAEINVAL
                            mstrLastErrorDesc = "WSAEINVAL: The Windows Sockets version specified by the application is not supported by this DLL"
    Case WSANOTINITIALISED
                            mstrLastErrorDesc = "WSANOTINITIALISED: A successful WSAStartup must occur before using this function"
    Case WSAENETDOWN
                            mstrLastErrorDesc = "WSAENETDOWN: The Windows Sockets implementation has detected that the network subsystem has failed"
    Case WSAEAFNOSUPPORT
                            mstrLastErrorDesc = "WSAEAFNOSUPPORT: The specified address family is not supported"
    Case WSAEINPROGRESS
                            mstrLastErrorDesc = "WSAEINPROGRESS: A blocking Windows Sockets operation is in progress"
    Case WSAEMFILE
                            mstrLastErrorDesc = "WSAEMFILE: No more file descriptors are available"
    Case WSAENOBUFS
                            mstrLastErrorDesc = "WSAENOBUFS: No buffer space is available. The socket cannot be created"
    Case WSAEPROTONOSUPPORT
                            mstrLastErrorDesc = "WSAEPROTONOSUPPORT: The specified protocol is not supported"
    Case WSAEPROTOTYPE
                            mstrLastErrorDesc = "WSAEPROTOTYPE: The specified protocol is the wrong type for this socket"
    Case WSAESOCKTNOSUPPORT
                            mstrLastErrorDesc = "WSAESOCKTNOSUPPORT: The specified socket type is not supported in this address family"
    Case WSAEADDRINUSE
                            mstrLastErrorDesc = "WSAEADDRINUSE: The specified address is already in use"
    Case WSAEINTR
                            mstrLastErrorDesc = "WSAEINTR: The (blocking) call was canceled using WSACancelBlockingCall"
    Case WSAEADDRNOTAVAIL
                            mstrLastErrorDesc = "WSAEADDRNOTAVAIL: The specified address is not available from the local computer"
    Case WSAECONNREFUSED
                            mstrLastErrorDesc = "WSAECONNREFUSED: The attempt to connect was forcefully rejected"
    Case WSAEFAULT
                            mstrLastErrorDesc = "WSAEFAULT: The namelen argument is incorrect"
    Case WSAEISCONN
                            mstrLastErrorDesc = "WSAEISCONN: The socket is already connected"
    Case WSAENETUNREACH
                            mstrLastErrorDesc = "WSAENETUNREACH: The network cant be reached from this host at this time"
    Case WSAENOTSOCK
                            mstrLastErrorDesc = "WSAENOTSOCK: The descriptor is not a socket"
    Case WSAETIMEDOUT
                            mstrLastErrorDesc = "WSAETIMEDOUT: Attempt to connect timed out without establishing a connection"
    Case WSAEWOULDBLOCK
                            mstrLastErrorDesc = "WSAEWOULDBLOCK: The socket is marked as nonblocking and the connection cannot be completed immediately. It is possible to select the socket while it is connecting by selecting it for writing"
    Case WSAEACCES
                            mstrLastErrorDesc = "The requested address is a broadcast address, but the appropriate flag was not set"
    Case WSAENETRESET
                            mstrLastErrorDesc = "The connection must be reset because the Windows Sockets implementation dropped it"
    Case WSAENOTCONN
                            mstrLastErrorDesc = "The socket is not connected"
    Case WSAEOPNOTSUPP
                            mstrLastErrorDesc = "MSG_OOB was specified, but the socket is not of type SOCK_STREAM"
    Case WSAESHUTDOWN
                            mstrLastErrorDesc = "The socket has been shutdown it is not possible to send on a socket after shutdown has been invoked with how set to 1 or 2"
    Case WSAEMSGSIZE
                            mstrLastErrorDesc = "The socket is of type SOCK_DGRAM, and the datagram is larger than the maximum supported by the Windows Sockets implementation"
    Case WSAECONNABORTED
                            mstrLastErrorDesc = "The virtual circuit was aborted due to timeout or other failure"
    Case WSAECONNRESET
                            mstrLastErrorDesc = "The virtual circuit was reset by the remote side"
    Case Else
         mstrLastErrorDesc = "Case unknown in TCPIP::SetLastErrorCode = " & mlngLastErrorCode
    End Select
    
    mstrLastErrorDesc = strAdditionalInfo & vbCrLf & mstrLastErrorDesc
End Sub

Public Property Let RemotePort(ByVal lngRemotePort As Long)
    mlngRemotePort = lngRemotePort
End Property

Public Property Let RemoteHostIP(ByVal strRemoteIP As String)
    mstrRemoteHostIP = strRemoteIP
End Property

Public Property Let LocalHostIP(ByVal strLocalHostIP As String)
    mstrLocalHostIP = strLocalHostIP
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = mstrLastErrorDesc
End Property

Public Property Get ErrorCode() As Long
    ErrorCode = mlngLastErrorCode
End Property

Private Sub Class_Terminate()
    If mlngSocket <> 0 Then ShutdownConnection
End Sub
