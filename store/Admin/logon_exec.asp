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

'*************************************************************************

'Logoff
if lCase(trim(request.QueryString("action"))) = "logoff" then 
	session.abandon 
	response.redirect "logon.asp" 
end if

'Check UserID and Password
if len(trim(Request.Form("adminUser"))) = 0 _
or len(trim(Request.Form("adminPass"))) = 0 then

	'Give error
	session.abandon
	Response.Redirect "logon.asp?msg=" & server.URLEncode("ERROR : Invalid UserID or Password.")

end if

'Check adminUser
if trim(Request.Form("adminUser")) = trim(adminUser) and trim(Request.Form("adminPass")) = trim(adminPass) then

	'Logon Administrator
	session(storeID & "adminLoggedOn") = "0"
	Response.Redirect "utilities.asp"

elseif trim(Request.Form("adminUser")) = trim(nonAdminUser) and trim(Request.Form("adminPass")) = trim(nonAdminPass) then

	'Logon Non Administrator
	session(storeID & "adminLoggedOn") = "1"
	Response.Redirect "utilities.asp"
	
else
	
	'Give error
	session.abandon
	Response.Redirect "logon.asp?msg=" & server.URLEncode("ERROR : Invalid UserID or Password.")
	
end if
%>
