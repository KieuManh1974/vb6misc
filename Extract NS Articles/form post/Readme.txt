			Form Poster Program
			===================
 Licence:
Do anything you want with it.
 Requirements:
The Microsoft Internet Transfer Control (MSINET.OCX).
 Contact:
If you want to send comments to me etc. do so at: bribobs@mail.com (not bribobs@hotmail.com)

The Microsoft Internet Transfer Control supports the use of the "POST" method, which allows programmers to submit filled in web forms to web servers. If you look at the source of a web form, you will find:
(1) A URL, embedded in html which looks something like this:
<FORM action="http://www.response-o-matic.com/cgi-bin/rom.pl" method="POST">

(2) One of more fields, of the form:
<INPUT TYPE="hidden" NAME="email_subject_line" VALUE="Response-O-matic Test">

To submit a response to a form, examine the html, and put the URL 
(in this case "http://www.response-o-matic.com/cgi-bin/rom.pl") in the TxtURL textbox (without the quotation marks)

In the TxtFields textboxes put on separate lines each field you wish to return with the field name and return value separated by a colon; in the above case you would put
email_subject_line:Response-O-matic Test
on a line in TxtURL. It is normally necessary to return all fields. Fields of all types, e.g. text boxes, check boxes, radio buttons can be returned.

The data POSTed to the server is of the form:
fieldname1=value1&fieldname2=value2&fieldname3=value3&fieldname4=value4... etc.

Some web forms require username and password, and the program allows these to be entered.
I have used this program to send text messages to my "Orange" mobile phone in the UK using their web interface at:
http://www.orange.co.uk/cgi-bin/register/sms.pl
and this requires registration, and a username and password.

Forms generate a response page, which is returned to the program and saved to disk, and then opened in the default browser.

To test the program the free service provided by www.response-o-matic.com may be used. Forms can be submitted to their server, and the response will be emailed to you.


Brian Reilly Dec 8th 2000