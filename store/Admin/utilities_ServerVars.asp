<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Display Server Variables
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
<%
'Work Fields
dim I

'*************************************************************************

'Are we in demo mode?
if demoMode = "Y" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("DEMO MODE. Sorry, this featured is NOT available in Demo Mode.")
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Display Server Variables</font></b>
	<br><br>
</P>

<table border="1" cellspacing="1" cellpadding="1" width="100%">
<%
'Display ServerVariables
for each I in Request.ServerVariables
	Response.Write "<tr><td><b>" & I & "</b>" & "</td><td>" & Request.ServerVariables(I) & "&nbsp;</td></tr>"
next
%>
</table>
	
<!--#include file="_INCfooter_.asp"-->
