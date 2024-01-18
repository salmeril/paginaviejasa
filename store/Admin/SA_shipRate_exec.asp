<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Shipping Rates Maintenance
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

'ShipRates
dim idShip
dim idShipMethod
dim locShipZone
dim unitType
dim unitsFrom
dim unitsTo
dim addAmt
dim addPerc

'ShipMethod
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

'Get idShip
if action = "edit" or action = "del" then

	idShip = trim(Request.Form("idShip"))
	if len(idShip) = 0 then
		idShip = trim(Request.QueryString("idShip"))
	end if
	if idShip = "" or not isNumeric(idShip) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Shipping Rate ID.")
	else
		idShip = CLng(idShip)
	end if
	
end if

if action = "edit" or action = "add" then

	'Get idShipMethod
	idShipMethod = trim(Request.Form("idShipMethod"))
	if len(idShipMethod) = 0 or not Isnumeric(idShipMethod) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Shipping Method.")
	end if
	
	'Check idShipMethod is valid
	mySQL = "SELECT idShipMethod " _
	      & "FROM   shipMethod " _
	      & "WHERE  idShipMethod = " & idShipMethod
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Shipping Method.")
	end if
	call closeRS(rs)
	
	'Get locShipZone
	locShipZone = trim(Request.Form("locShipZone"))
	if len(locShipZone) = 0 or not Isnumeric(locShipZone) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Shipping Zone.")
	end if
	
	'Get unitType
	unitType = trim(Request.Form("unitType"))
	if unitType <> "P" and unitType <> "W" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Rate Type.")
	end if
	
	'Get unitsFrom
	unitsFrom = trim(Request.Form("unitsFrom"))
	if len(unitsFrom) = 0 or not Isnumeric(unitsFrom) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Price/Weight FROM value.")
	end if
	unitsFrom = CDbl(unitsFrom)
	
	'Get unitsTo
	unitsTo = trim(Request.Form("unitsTo"))
	if len(unitsTo) = 0 or not Isnumeric(unitsTo) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Price/Weight TO value.")
	end if
	unitsTo = CDbl(unitsTo)
	
	'Check that unitsTo > unitsFrom
	if unitsTo <= unitsFrom then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Price/Weight TO value must be greater than FROM value.")
	end if

	'Get addAmt
	addAmt = trim(Request.Form("addAmt"))
	if len(addAmt) > 0 then
		if not Isnumeric(addAmt) then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Add Amount.")
		end if
	end if
	
	'Get addPerc
	addPerc = trim(Request.Form("addPerc"))
	if len(addPerc) > 0 then
		if not Isnumeric(addPerc) then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Add Percentage.")
		end if
	end if
	
	'Check that addAmt or addPerc is entered
	if len(addAmt) = 0 and len(addPerc) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Add Amount and/or Add Percentage must have a value.")
	end if
	
	'Assign null values to addAmt or addPerc if applicable
	if len(addAmt) = 0 then
		addAmt = "null"
	else
		addAmt = CDbl(addAmt)
	end if
	if len(addPerc) = 0 then
		addPerc = "null"
	else
		addPerc = CDbl(addPerc)
	end if
	
end if

'ADD
if action = "add" then

	'Add Shipping Rate
	mySQL = "INSERT INTO shipRates (" _
		  & "idShipMethod,locShipZone,unitType," _
		  & "unitsFrom,unitsTo,addAmt,addPerc" _
	      & ") VALUES (" _
	      &       idShipMethod	& " ," _
	      &       locShipZone	& " ," _
	      & "'" & unitType		& "'," _
	      &       unitsFrom		& " ," _
	      &       unitsTo		& " ," _
	      &       addAmt		& " ," _
	      &       addPerc		& " " _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	call closedb()
	Response.Redirect "SA_shipRate.asp?recallCookie=1&msg=" & server.URLEncode("Shipping Rate record was Added.")
	
end if

'DELETE or BULK DELETE
if action = "del" or action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of idShip that will be deleted

	'If just one delete is being performed, we populate just the 
	'first position in the delete array, else we populate the array
	'with a list of all the records that were selected for deletion.
	if action = "del" then
		delArray = split(idShip)
	else
		delArray = split(Request.Form("idShip"),",")
	end if
	
	'Loop through list of records and delete one by one
	for delI = LBound(delArray) to UBound(delArray)

		'Delete record from shipRates
		mySQL = "DELETE FROM shipRates " _
		      & "WHERE idShip = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)
		
	next

	call closedb()
	Response.Redirect "SA_shipRate.asp?recallCookie=1&msg=" & server.URLEncode("Shipping Rate records(s) were Deleted.")

end if

'EDIT
if action = "edit" then

	'Update Record
	mySQL = "UPDATE shipRates SET " _
	      & "idShipMethod=" & idShipMethod	& "," _
	      & "locShipZone="	& locShipZone	& "," _
	      & "unitType='"	& unitType		& "'," _
	      & "unitsFrom="	& unitsFrom		& "," _
	      & "unitsTo="		& unitsTo		& "," _
	      & "addAmt="		& addAmt		& "," _
	      & "addPerc="		& addPerc		& " " _
	      & "WHERE  idShip = " & idShip
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_shipRate.asp?recallCookie=1&msg=" & server.URLEncode("Shipping Rate record was Updated.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_shipRate.asp?recallCookie=1"

%>
