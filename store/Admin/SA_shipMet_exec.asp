<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Shipping Method Maintenance
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
dim mySQL, cn, rs

'ShipMethod
dim idShipMethod
dim shipDesc
dim status

'Work Fields
dim action

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
if  action <> "edit" _
and action <> "del" _
and action <> "add" _
and action <> "bulkdel" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idShipMethod
if action = "edit" or action = "del" then

	idShipMethod = trim(Request.Form("idShipMethod"))
	if len(idShipMethod) = 0 then
		idShipMethod = trim(Request.QueryString("idShipMethod"))
	end if
	if idShipMethod = "" or not isNumeric(idShipMethod) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Shipping Method ID.")
	else
		idShipMethod = CLng(idShipMethod)
	end if

end if

if action = "edit" or action = "add" then

	'Get Description
	shipDesc = trim(Request.Form("shipDesc"))
	shipDesc = replace(shipDesc,"""","") 'To prevent HTML field terminations
	if len(shipDesc) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Description.")
	end if

	'Get Status
	status = trim(Request.Form("status"))
	if status <> "A" then
		status = "I"
	end if

	'Check no Description duplicates
	mySQL = "SELECT shipDesc " _
	      & "FROM   shipMethod " _
	      & "WHERE  shipDesc = '" & replace(shipDesc,"'","''") & "' "
	if action = "edit" then
		mySQL = mySQL  & "AND idShipMethod <> " & idShipMethod & " "
	end if
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("A Shipping Method with that Description already exists.")
	end if
	call closeRS(rs)
	
end if

'ADD
if action = "add" then

	'Add Option
	mySQL = "INSERT INTO ShipMethod (" _
	      & "shipDesc,status" _
	      & ") VALUES (" _
	      & "'" & replace(shipDesc,"'","''") & "','" & status & "'"  _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	'Get idShipMethod of INSERTed Record
	mySQL = "SELECT MAX(idShipMethod) AS maxIdShipMethod " _
		  & "FROM   ShipMethod "
	set rs = openRSexecute(mySQL)
	idShipMethod = rs("maxIdShipMethod")
	call closeRS(rs)
	
	call closedb()
	Response.Redirect "SA_shipMet.asp?msg=" & server.URLEncode("Shipping Method was Added.")
	
end if

'DELETE or BULK DELETE
if action = "del" or action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of idShipMethod that will be deleted

	'If just one delete is being performed, we populate just the 
	'first position in the delete array, else we populate the array
	'with a list of all the records that were selected for deletion.
	if action = "del" then
		delArray = split(idShipMethod)
	else
		delArray = split(Request.Form("idShipMethod"),",")
	end if

	'Set CursorLocation of the Connection Object to Client
	cn.CursorLocation = adUseClient
	
	'Loop through list of records and delete one by one
	for delI = LBound(delArray) to UBound(delArray)
	
		'BEGIN Transaction
		cn.BeginTrans
	
		'Delete records from ShipMethod
		mySQL = "DELETE FROM ShipMethod " _
		      & "WHERE idShipMethod = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'Delete records from ShipRates
		mySQL = "DELETE FROM ShipRates " _
		      & "WHERE idShipMethod = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'END Transaction
		cn.CommitTrans
		
	next

	call closedb()
	Response.Redirect "SA_shipMet.asp?msg=" & server.URLEncode("Shipping Method(s) were Deleted.")

end if

'EDIT
if action = "edit" then

	'Update Record
	mySQL = "UPDATE ShipMethod SET " _
	      & "       shipDesc = '" & replace(shipDesc,"'","''") & "', " _
	      & "       status   = '"  & status & "' " _ 
	      & "WHERE  idShipMethod = " & idShipMethod
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_shipMet.asp?msg=" & server.URLEncode("Shipping Method was Updated.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_shipMet.asp"

%>
