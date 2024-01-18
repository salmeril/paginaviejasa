<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Script display errors and general messages
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
%>
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="_INCconfig_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="_INCappFunctions_.asp"-->
<%
'Database
dim connTemp
dim errMsg, msg, returnURL

'Session
dim idOrder
dim idCust

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	call errorDB(langErrConfig,"")
end if

'Get/Set Cart/Order Session
idOrder = sessionCart()

'Get/Set Customer Session
idCust  = sessionCust()

'Get input parms
msg       = server.HTMLEncode(trim(Request.QueryString("msg")))
errMsg    = server.HTMLEncode(trim(Request.QueryString("errMsg")))
returnURL = server.HTMLEncode(trim(Request.QueryString("returnURL")))
%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<table width="100%" border="0" cellspacing="0" cellpadding="2">
	<tr><td valign=middle class="CPpageHead">
		<b><%=langGenSystemMessage%></b><br>
	</td></tr>
</table>

<br><br>

<table width="100%" border="0" cellspacing="0" cellpadding="5">
<tr><td>

	<font size=2>
<%
		if len(errmsg) > 0 then
			Response.Write "<font color=red>" & errMsg & "</font> "
		else
			Response.Write msg & " "
		end if
		if len(returnURL) > 0 then
			Response.Write "<a href=""" & returnURL & """>" & langGenClickToReturn & "</a>"
		end if
%>
	</font>
	
</td></tr>
</table>

<br><br><br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%
call closeDB()
%>