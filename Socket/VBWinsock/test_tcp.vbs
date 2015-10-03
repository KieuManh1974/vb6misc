'This script test the TCPIP component.
'It connects to Microsoft Web Site
'to retrieve its home page.

Dim tcp, strData, l, bOk

Set tcp = WScript.CreateObject("VBWinsock.TCPIP")
tcp.LocalHostIP = "168.234.135.22"
tcp.RemoteHostIP = "207.46.130.150"
tcp.RemotePort = 80

tcp.OpenConnection
tcp.SendData("GET http://www.microsoft.com/" & Chr(13) & Chr(10))

bOk = tcp.ReceiveData(strData, l)  'Only one call to RecvData might not
				   ' be enough to retrieve all the page.
WScript.Echo strData		

tcp.ShutdownConnection
Set tcp = Nothing