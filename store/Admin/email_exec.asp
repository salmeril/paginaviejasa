<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : send Email
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
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="../Scripts/_INCappEmail_.asp"-->
<%
'Email
dim emailFrom
dim emailTo
dim emailSubj
dim emailBody
dim contType

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

'Get Form Fields
emailFrom   = trim(request.form("emailFrom"))
emailTo     = trim(request.form("emailTo"))
emailSubj   = trim(request.form("emailSubj"))
emailBody   = trim(request.form("emailBody"))
contType    = trim(request.form("contType"))

'Do some checks
if instr(emailFrom,"@")=0 or instr(emailFrom,".")=0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Email FROM address.")
end if
if instr(emailTo,"@")=0 or instr(emailTo,".")=0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Email TO address.")
end if
if len(emailSubj) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Email subject.")
end if
if len(emailBody) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Email message body.")
end if
if isNumeric(contType) then
	contType = CLng(contType)
else
	contType = 0
end if

'Send Email
on error resume next
call sendmail (pCompany, emailFrom, emailTo, emailSubj, emailBody, contType)
if err.number <> 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("An error occurred." & err.Description)
end if

'Return to form
response.redirect "email.asp?msg=" & server.URLEncode("Your Email was sent.")
%>
