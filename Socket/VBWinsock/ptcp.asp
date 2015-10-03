<%@ LANGUAGE="VBSCRIPT" %>
<% 

Dim bOk, strSourceCode, strErrorDesc, tcp, l

Server.ScriptTimeOut = 180

Set tcp = Server.CreateObject("VBWinsock.TCPIP")
tcp.LocalHostIP = Request.Form("txtLocalHostIP")
tcp.RemoteHostIP = Request.Form("txtRemoteHostIP")
tcp.RemotePort = 80

bOk = tcp.OpenConnection
If bOk Then bOk = tcp.SendData("GET " & Request.Form("txtURL") & Chr(13) & Chr(10))

'Only one call to RecvData might not
'be enough to retrieve all the page.
If bOk Then bOk = tcp.ReceiveData(strSourceCode, l)  
				   
If NOT bOk Then strErrorDesc = tcp.ErrorDescription

tcp.ShutdownConnection
Set tcp = Nothing	

%>

<html>

<head>
<title>Results</title>
</head>

<body bgcolor="#FFFFFF">

<p>Here is the source code of</p>

<p><a href="<%=Request.Form("txtURL")%>"><%=Request.Form("txtURL")%></a></p>

<% If bOk Then %>
   <p><textarea rows="21" name="S1" cols="51"><%=strSourceCode%></textarea></p>
   <p>Remember that is possible that the page isn't complete since we only called ReceiveData() once...</p>
<% Else 
   Response.Write strErrorDesc
End If%>

</body>
</html>
