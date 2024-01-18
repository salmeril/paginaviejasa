<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Check Email
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
Option explicit
Response.Buffer = true
const adminLevel = 0
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="../Scripts/_INCappEmail_.asp"-->
<%
'Database
dim mySQL, cn, rs

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Close Database Connection
call closedb()

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Check Email</font></b>
	<br><br>
</P>

<%
'Send Test Email
call sendmail (pCompany, pEmailAdmin, pEmailAdmin, "Email Test", "Congratulations. Your email is working!", 0)
%>

<table border=0 cellspacing=0 cellpadding=10 class="textBlock">
<tr><td>
	<font size=2>
		An Email was sent to :
		<br><br>
		<b><font color=red><%=pEmailAdmin%></font></b>
		<br><br>
		Check your Email Inbox to see if this message was delivered. If 
		it was delivered successfully then your Email is set up properly. 
		If you don't get an email, make sure that your Email settings 
		in your configuration is entered correctly. You may need 
		to check with your Web Hosting company for more information 
		regarding the setup of your Email server.
	</font>
	<br><br>
</td></tr>
</table>

<!--#include file="_INCfooter_.asp"-->
