<% if request("cmdSend") <> "" then
set objSendMail = Server.CreateObject("vbSendMail.clsSendMail")

objSendMail.SMTPHost = Cstr(Request.Form("SMTPServer"))
objSendMail.From = Cstr(Request.Form("FromEmail"))                        
objSendMail.FromDisplayName = Cstr(Request.Form("FromName"))   
objSendMail.Recipient = Cstr(Request.Form("ToEmail"))                  
objSendMail.RecipientDisplayName = Cstr(Request.Form("ToName"))           
objSendMail.ReplyToAddress = Cstr(Request.Form("FromEmail"))               
objSendMail.Subject = Cstr(Request.Form("Subject"))           
objSendMail.Message = Cstr(Request.Form("Message"))                  
if Request("AsHTML") <> "" then objSendMail.AsHTML = True


objSendMail.Send

Set objSendMail = Nothing                            
sResults = "<FONT FACE = 'ARIAL' COLOR = 'RED'>Mail has been sent</FONT><P>"
End if
%>

<HTML>
<HEAD>
<TITLE>vbSendMail ASP Example</TITLE>
</HEAD>
<BODY>
<% response.write sResults %>
<CENTER><FONT SIZE = +1 FACE = 'ARIAL'><B>vbSendMail ASP example</B></CENTER><P>
<FORM ACTION = "vbSendMail.asp" METHOD = "POST">
<TABLE>
<TR>
<TD>SMTP Server:</TD>
<TD><INPUT NAME = SMTPServer TYPE=TEXT SIZE=30></TD>
</TR>

<TR>
<TD>From (Name):</TD>
<TD><INPUT NAME = FromName TYPE = TEXT SIZE=30></TD>
</TR>

<TR>
<TD>From (E-mail):</TD>
<TD><INPUT NAME= FromEmail TYPE=TEXT SIZE=30></TD>
</TR>

<TR>
<TD>To (Name):</TD>
<TD><INPUT NAME = ToName TYPE = TEXT SIZE=30></TD>
</TR>

<TR>
<TD>To (Email):</TD>
<TD><INPUT NAME = ToEmail TYPE = TEXT SIZE=30></TD>
</TR>

<TR>
<TD>Subject :</TD>
<TD><INPUT NAME = Subject TYPE = TEXT SIZE=30></TD>
</TR>

<TR>
<TD VALIGN = TOP>MESSAGE:</TD>
<TD VALIGN = TOP><TEXTAREA NAME=MESSAGE ROWS=8 COLS = 40></TEXTAREA></TD></TR>
<TR>
<TD>&nbsp;</TD><TD><INPUT TYPE = CHECKBOX NAME = AsHTML></INPUT>Message is in HTML format</TD>
</TR>
</TABLE><P>

<CENTER>
<INPUT TYPE = SUBMIT NAME="cmdSend" VALUE = "Send Message">
</CENTER></FONT>
</FORM>


</BODY>
</HTML>

