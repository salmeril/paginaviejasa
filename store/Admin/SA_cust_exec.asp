<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Customer Maintenance
' Product  : CandyPress eCommerce Administration
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
<%

'Database
dim mySQL, cn, rs, rs2

'Customer
dim idCust
dim status
dim dateCreated
dim dateCreatedInt
dim name
dim lastName
dim customerCompany
dim phone
dim email
dim password
dim address
dim city
dim locState
dim locCountry
dim zip
dim paymentType
dim shippingName
dim shippingLastName
dim shippingPhone
dim shippingAddress
dim shippingCity
dim shippingLocState
dim shippingLocCountry
dim shippingZip
dim futureMail
dim generalComments
dim taxExempt

'Work Fields
dim action
dim orderCount

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Get action
action = trim(lCase(Request.Form("action")))
if len(action) = 0 then
	action = trim(lCase(Request.QueryString("action")))
end if
if action <> "edit" and action <> "del" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idCust
if action = "edit" or action = "del" then

	idCust = trim(Request.Form("idCust"))
	if len(idCust) = 0 then
		idCust = trim(Request.QueryString("idCust"))
	end if
	if idCust = "" or not isNumeric(idCust) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Customer ID.")
	else
		idCust = CLng(idCust)
	end if
	
end if

'EDIT
if action = "edit" then

	'Get fields 
	status			= trim(Request.Form("status"))
	taxExempt		= trim(Request.Form("taxExempt"))
	Email			= trim(replace(Request.Form("Email"),"""",""))
	generalComments	= trim(replace(Request.Form("generalComments"),"""",""))
	
	'Update record
	mySQL="UPDATE customer SET " _
		& "status='"	& status					& "'," _
		& "taxExempt='"	& taxExempt					& "'," _
		& "Email='"		& replace(Email,"'","''")	& "'," _
		& "generalComments='"		& replace(generalComments,"'","''")		& "'" _
		& "WHERE idCust = " & idCust
	set rs = openRSexecute(mySQL)
	
	call closedb()
	Response.Redirect "SA_cust.asp?recallCookie=1&msg=" & server.URLEncode("Customer was Updated.")

end if

'DELETE
if action = "del" then

	'Check if there are any Orders for this Customer
	mySQL = "SELECT COUNT(*) AS orderCount " _
		  & "FROM   cartHead " _
		  & "WHERE  idCust=" & idCust
	set rs = openRSexecute(mySQL)
	orderCount = rs("orderCount")
	call closeRS(rs)
	
	'If there were orders, then this Customer can not be deleted
	if orderCount > 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Customer can not be deleted because there are Orders linked to it.")
	end if

	'Set CursorLocation of the Connection Object to Client
	cn.CursorLocation = adUseClient
	
	'BEGIN Transaction
	cn.BeginTrans
	
	'Delete records from customer
	mySQL = "DELETE FROM customer " _
	      & "WHERE idCust = " & idCust
	set rs = openRSexecute(mySQL)

	'END Transaction
	cn.CommitTrans

	call closedb()
	Response.Redirect "SA_cust.asp?recallCookie=1&msg=" & server.URLEncode("Customer was Deleted.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_cust.asp?recallCookie=1"

%>
