<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Fix Physical Structure of Database
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
<%
'Declare variables
dim mySQL, cn, rs
dim i
dim fixSQL
dim SQLArr

'*************************************************************************

'Open Database Connection
call openDb()

'Check to see if we need to apply fixes
fixSQL = trim(Request.Form("fixSQL"))

if len(fixSQL) > 0 then
	SQLArr = split(fixSQL,"*|*") 'If there were more than SQL statement
	for i = 0 to Ubound(SQLArr)
		if len(trim(SQLArr(i))) > 0 then
			set rs = openRSexecute(SQLArr(i))
		end if
	next
end if

'Close Database
call closedb()

Response.Redirect "utilities_DBstruc.asp"

%>