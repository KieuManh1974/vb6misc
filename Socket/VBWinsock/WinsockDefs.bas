Attribute VB_Name = "WinsockDefs"
Option Explicit

Public Const WSANOERROR = 0

'Address families.

Public Const AF_UNSPEC = 0        '               { unspecified }
Public Const AF_UNIX = 1          '               { local to host (pipes, portals) }
Public Const AF_INET = 2          '               { internetwork: UDP, TCP, etc. }
Public Const AF_IMPLINK = 3       '               { arpanet imp addresses }
Public Const AF_PUP = 4           '               { pup protocols: e.g. BSP }
Public Const AF_CHAOS = 5         '               { mit CHAOS protocols }
Public Const AF_IPX = 6           '               { IPX and SPX }
Public Const AF_NS = 6            '               { XEROX NS protocols }
Public Const AF_ISO = 7           '               { ISO protocols }
Public Const AF_OSI = AF_ISO      '               { OSI is ISO }
Public Const AF_ECMA = 8          '               { european computer manufacturers }
Public Const AF_DATAKIT = 9       '               { datakit protocols }
Public Const AF_CCITT = 10        '               { CCITT protocols, X.25 etc }
Public Const AF_SNA = 11          '               { IBM SNA }
Public Const AF_DECnet = 12       '               { DECnet }
Public Const AF_DLI = 13          '               { Direct data link interface }
Public Const AF_LAT = 14          '               { LAT }
Public Const AF_HYLINK = 15       '               { NSC Hyperchannel }
Public Const AF_APPLETALK = 16    '               { AppleTalk }
Public Const AF_NETBIOS = 17      '               { NetBios-style addresses }
Public Const AF_VOICEVIEW = 18    '               { VoiceView }
Public Const AF_FIREFOX = 19      '               { FireFox }
Public Const AF_UNKNOWN1 = 20     '               { Somebody is using this! }
Public Const AF_BAN = 21          '               { Banyan }
Public Const AF_MAX = 22          '

Public Const PF_UNSPEC = AF_UNSPEC
Public Const PF_UNIX = AF_UNIX
Public Const PF_INET = AF_INET
Public Const PF_IMPLINK = AF_IMPLINK
Public Const PF_PUP = AF_PUP
Public Const PF_CHAOS = AF_CHAOS
Public Const PF_NS = AF_NS
Public Const PF_IPX = AF_IPX
Public Const PF_ISO = AF_ISO
Public Const PF_OSI = AF_OSI
Public Const PF_ECMA = AF_ECMA
Public Const PF_DATAKIT = AF_DATAKIT
Public Const PF_CCITT = AF_CCITT
Public Const PF_SNA = AF_SNA
Public Const PF_DECnet = AF_DECnet
Public Const PF_DLI = AF_DLI
Public Const PF_LAT = AF_LAT
Public Const PF_HYLINK = AF_HYLINK
Public Const PF_APPLETALK = AF_APPLETALK
Public Const PF_VOICEVIEW = AF_VOICEVIEW
Public Const PF_FIREFOX = AF_FIREFOX
Public Const PF_UNKNOWN1 = AF_UNKNOWN1
Public Const PF_BAN = AF_BAN
Public Const PF_MAX = AF_MAX

Public Const SOCK_STREAM = 1&
Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYS_STATUS_LEN = 128

Public Const INVALID_SOCKET = -1&
Public Const SOCKET_ERROR = -1&
Public Const INADDR_NONE = &HFFFFFFFF

' All Windows Sockets error constants are biased by WSABASEERR from the "normal"

Public Const WSABASEERR = 10000

' Windows Sockets definitions of regular Microsoft C error constants

Public Const WSAEINTR = (WSABASEERR + 4)
Public Const WSAEBADF = (WSABASEERR + 9)
Public Const WSAEACCES = (WSABASEERR + 13)
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEMFILE = (WSABASEERR + 24)

' Windows Sockets definitions of regular Berkeley error constants

Public Const WSAEWOULDBLOCK = (WSABASEERR + 35)
Public Const WSAEINPROGRESS = (WSABASEERR + 36)
Public Const WSAEALREADY = (WSABASEERR + 37)
Public Const WSAENOTSOCK = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ = (WSABASEERR + 39)
Public Const WSAEMSGSIZE = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT = (WSABASEERR + 47)
Public Const WSAEADDRINUSE = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL = (WSABASEERR + 49)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSAENETUNREACH = (WSABASEERR + 51)
Public Const WSAENETRESET = (WSABASEERR + 52)
Public Const WSAECONNABORTED = (WSABASEERR + 53)
Public Const WSAECONNRESET = (WSABASEERR + 54)
Public Const WSAENOBUFS = (WSABASEERR + 55)
Public Const WSAEISCONN = (WSABASEERR + 56)
Public Const WSAENOTCONN = (WSABASEERR + 57)
Public Const WSAESHUTDOWN = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS = (WSABASEERR + 59)
Public Const WSAETIMEDOUT = (WSABASEERR + 60)
Public Const WSAECONNREFUSED = (WSABASEERR + 61)
Public Const WSAELOOP = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH = (WSABASEERR + 65)
Public Const WSAENOTEMPTY = (WSABASEERR + 66)
Public Const WSAEPROCLIM = (WSABASEERR + 67)
Public Const WSAEUSERS = (WSABASEERR + 68)
Public Const WSAEDQUOT = (WSABASEERR + 69)
Public Const WSAESTALE = (WSABASEERR + 70)
Public Const WSAEREMOTE = (WSABASEERR + 71)

Public Const WSAEDISCON = (WSABASEERR + 101)

' Extended Windows Sockets error constant definitions

Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)

' Error return codes from gethostbyname() and gethostbyaddr()
'  (when using the resolver). Note that these errors are
'  retrieved via WSAGetLastError() and must therefore follow
'  the rules for avoiding clashes with error numbers from
'  specific implementations or language run-time systems.
'  For this reason the codes are based at WSABASEERR+1001.
'  Note also that [WSA]NO_ADDRESS is defined only for
'  compatibility purposes.

' Authoritative Answer: Host not found

Public Const WSAHOST_NOT_FOUND = (WSABASEERR + 1001)
Public Const HOST_NOT_FOUND = WSAHOST_NOT_FOUND

' Non-Authoritative: Host not found, or SERVERFAIL

Public Const WSATRY_AGAIN = (WSABASEERR + 1002)
Public Const TRY_AGAIN = WSATRY_AGAIN

' Non recoverable errors, FORMERR, REFUSED, NOTIMP

Public Const WSANO_RECOVERY = (WSABASEERR + 1003)
Public Const NO_RECOVERY = WSANO_RECOVERY

' Valid name, no data record of requested type

Public Const WSANO_DATA = (WSABASEERR + 1004)
Public Const NO_DATA = WSANO_DATA

' no address, look for MX record

Public Const WSANO_ADDRESS = WSANO_DATA
Public Const NO_ADDRESS = WSANO_ADDRESS

' Windows Sockets errors redefined as regular Berkeley error constants.
' These are commented out in Windows NT to avoid conflicts with errno.h.
' Use the WSA constants instead.

Public Const EWOULDBLOCK = WSAEWOULDBLOCK
Public Const EINPROGRESS = WSAEINPROGRESS
Public Const EALREADY = WSAEALREADY
Public Const ENOTSOCK = WSAENOTSOCK
Public Const EDESTADDRREQ = WSAEDESTADDRREQ
Public Const EMSGSIZE = WSAEMSGSIZE
Public Const EPROTOTYPE = WSAEPROTOTYPE
Public Const ENOPROTOOPT = WSAENOPROTOOPT
Public Const EPROTONOSUPPORT = WSAEPROTONOSUPPORT
Public Const ESOCKTNOSUPPORT = WSAESOCKTNOSUPPORT
Public Const EOPNOTSUPP = WSAEOPNOTSUPP
Public Const EPFNOSUPPORT = WSAEPFNOSUPPORT
Public Const EAFNOSUPPORT = WSAEAFNOSUPPORT
Public Const EADDRINUSE = WSAEADDRINUSE
Public Const EADDRNOTAVAIL = WSAEADDRNOTAVAIL
Public Const ENETDOWN = WSAENETDOWN
Public Const ENETUNREACH = WSAENETUNREACH
Public Const ENETRESET = WSAENETRESET
Public Const ECONNABORTED = WSAECONNABORTED
Public Const ECONNRESET = WSAECONNRESET
Public Const ENOBUFS = WSAENOBUFS
Public Const EISCONN = WSAEISCONN
Public Const ENOTCONN = WSAENOTCONN
Public Const ESHUTDOWN = WSAESHUTDOWN
Public Const ETOOMANYREFS = WSAETOOMANYREFS
Public Const ETIMEDOUT = WSAETIMEDOUT
Public Const ECONNREFUSED = WSAECONNREFUSED
Public Const ELOOP = WSAELOOP
Public Const ENAMETOOLONG = WSAENAMETOOLONG
Public Const EHOSTDOWN = WSAEHOSTDOWN
Public Const EHOSTUNREACH = WSAEHOSTUNREACH
Public Const ENOTEMPTY = WSAENOTEMPTY
Public Const EPROCLIM = WSAEPROCLIM
Public Const EUSERS = WSAEUSERS
Public Const EDQUOT = WSAEDQUOT
Public Const ESTALE = WSAESTALE
Public Const EREMOTE = WSAEREMOTE

Type WSA_DATA
    wVersion As Integer
    wHighVersion As Integer
    strDescription(WSADESCRIPTION_LEN + 1) As Byte
    strSystemStatus(WSASYS_STATUS_LEN + 1) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Type IN_ADDR
    S_addr As Long
End Type

Type SOCK_ADDR
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR
    sin_zero(0 To 7) As Byte
End Type

Public Const FD_SETSIZE = 64
Type FD_SET
    fd_count As Long
    fd_array(0 To FD_SETSIZE - 1) As Long
End Type

Type TIME_VAL
    tv_sec As Long
    tv_usec As Long
End Type

Declare Function bind Lib "wsock32" (ByVal s As Long, addr As SOCK_ADDR, ByVal namelen As Long) As Long
Declare Function closesocket Lib "wsock32" (ByVal s As Long) As Long
Declare Function connect Lib "wsock32" (ByVal s As Long, name As SOCK_ADDR, ByVal namelen As Integer) As Long
Declare Function inet_addr Lib "wsock32" (ByVal cp As String) As Long
Declare Function htons Lib "wsock32" (ByVal hostshort As Integer) As Integer
Declare Function recv Lib "wsock32" (ByVal s As Long, buffer As Any, ByVal length As Long, ByVal flags As Long) As Long
Declare Function send Lib "wsock32" (ByVal s As Long, buffer As Any, ByVal length As Long, ByVal flags As Long) As Long
Declare Function shutdown Lib "wsock32" (ByVal s As Long, ByVal how As Long) As Long
Declare Function sselect Lib "wsock32" Alias "select" (ByVal nfds As Long, readfds As FD_SET, writefds As FD_SET, exceptfds As FD_SET, timeout As TIME_VAL) As Long
Declare Function socket Lib "wsock32" (ByVal af As Long, ByVal type_specification As Long, ByVal protocol As Long) As Long
Declare Function WSACancelBlockingCall Lib "wsock32" () As Long
Declare Function WSACleanup Lib "wsock32" () As Long
Declare Function WSAGetLastError Lib "wsock32" () As Long
Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Integer, wsData As WSA_DATA) As Long

