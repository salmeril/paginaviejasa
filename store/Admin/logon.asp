<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Logon
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
const adminLevel = 1
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<%

'Demo Fields
dim demoUser
dim demoPass

'*************************************************************************

%>
<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Administration Log On</font></b>
	<br><br>
</P>

<%
if len(trim(Request.QueryString("msg"))) > 0 then
%>
	<font color=red><%=Request.QueryString("msg")%></font>
	<br><br>
<%
else
%>
	<br><br>
<%
end if

'Check if we are running in Demo Mode
if demoMode = "Y" then
	demoUser = adminUser
	demoPass = adminPass
else
	demoUSer = ""
	demoPass = ""
end if
%>

<TABLE BORDER="0" CELLPADDING="0" cellspacing="1" width=350>
	<TR> 
		<form METHOD="POST" name="Logonform" action="logon_exec.asp">
		<TD align=left valign=top width="50%" nowrap>
			Admin User ID :<br>
			<input type=text name=adminUser size=20 value="<%=demoUser%>">
			<br><br>
			Admin Password :<br>
			<input type=password name=adminPass size=20 value="<%=demoPass%>">
			<br><br>
			<input type="submit" name="Submit" value="Log On">
			<br><br>
		</TD>
		</form>
    </TR>
</TABLE>

<br><br>

<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
<tr><td>
	<b>Note : </b> The Administration utility is visually optimised for 
	use in <b>Internet Explorer 5.5</b> (or later versions) or <b>Netscape 6.0</b> 
	(or later versions) browsers, although it will still work with 
	version 4.0 browsers. Please ensure that you have Cookies 
	and JavaScript enabled.
<%
	if LCase(Request.ServerVariables("HTTPS")) <> "on" then
%>
		<br><br>
		<b>Note : </b> It appears that you are running the Administration 
		utility in an <font color=red>UNSECURE</font> session. Though not 
		technically required, you are advised to switch to a secure session 
		if you will be viewing sensitive information such as Credit Cards, 
		etc.
<%
	end if
%>
</td></tr>
</table>

<br><br>

<!--#include file="_INCfooter_.asp"--> 
