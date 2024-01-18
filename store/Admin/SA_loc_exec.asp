<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Location Maintenance
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

'Locations
dim idLocation
dim locName
dim locCountry
dim locState
dim locTax
dim locShipZone
dim locStatus

'Work Fields
dim action
dim locCountryOld

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
and action <> "editstate" _
and action <> "addstate" _
and action <> "bulkdel" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idLocation
if action = "edit" or action = "editstate" or action = "del" then

	idLocation = trim(Request.Form("idLocation"))
	if len(idLocation) = 0 then
		idLocation = trim(Request.QueryString("idLocation"))
	end if
	if idLocation = "" or not isNumeric(idLocation) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	else
		idLocation = CLng(idLocation)
	end if
	
end if

'Get locCountry, locState from DB
if action = "del" then

	mySQL="SELECT locCountry,locState " _
	    & "FROM   Locations " _
	    & "WHERE  idLocation = " & idLocation & " "
	set rs = openRSexecute(mySQL)
	if not rs.EOF then
		locCountry = rs("locCountry")
		locState   = rs("locState")
	else
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if
	call closeRS(rs)
		
end if

'Get other info from FORM
if action = "edit" or action = "editstate" or action = "add" or action = "addstate" then

	'Get locCountry
	locCountry = trim(Request.Form("locCountry"))
	if len(locCountry) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Country Code.")
	end if
	
	'Check for duplicate Country Codes
	if action = "edit" or action = "add" then
		mySQL="SELECT COUNT(*) AS recCount " _
		    & "FROM   Locations " _
		    & "WHERE  locCountry = '" & locCountry & "' "
		if action = "edit" then
			mySQL = mySQL _
				 & "AND (locState IS NULL OR locState = '') " _
				 & "AND  idLocation <> " & idLocation
		end if
		set rs = openRSexecute(mySQL)
		if rs("recCount") <> 0 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Country Code already exists.")
		end if
		call closeRS(rs)
	end if
	
	'Get locState
	if action = "editstate" or action = "addstate" then
		locState = trim(Request.Form("locState"))
		if len(locState) = 0 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid State/Province Code.")
		end if
	else
		locState = ""
	end if
	
	'Check for duplicate State Codes
	if action = "editstate" or action = "addstate" then
		mySQL="SELECT COUNT(*) AS recCount " _
		    & "FROM   Locations " _
		    & "WHERE  locCountry = '" & locCountry & "' " _
		    & "AND    locState   = '" & locState   & "' "
		if action = "editstate" then
			mySQL = mySQL & "AND idLocation <> " & idLocation
		end if
		set rs = openRSexecute(mySQL)
		if rs("recCount") <> 0 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("State/Province Code already exists.")
		end if
		call closeRS(rs)
	end if

	'Get Country/State Name
	locName = trim(Request.Form("locName"))
	if len(locName) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Name.")
	end if
	
	'Get Tax Rate
	locTax = trim(Request.Form("locTax"))
	if len(locTax) = 0 then
		locTax = 0.00
	end if
	if not Isnumeric(locTax) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Tax Rate.")
	else
		locTax = CDbl(locTax)
	end if
	
	'Get Shipping Zone
	locShipZone = trim(Request.Form("locShipZone"))
	if len(locShipZone) = 0 or not Isnumeric(locShipZone) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Shipping Zone.")
	end if
	
	'Get Status
	locStatus = UCase(trim(Request.Form("locStatus")))
	if locStatus <> "A" and locStatus <> "I" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Status.")
	end if

	'Make sure there are no Double Quotes in Location Name
	locName	= replace(locName,"""","")

end if

'ADD - Add Country
if action = "add" then

	'Add Country 
	mySQL = "INSERT INTO Locations (" _
		  & "locName,locCountry,locState,locTax,locShipZone,locStatus" _
	      & ") VALUES (" _
	      & "'" & replace(locName,"'","''") & "'," _
	      & "'" & locCountry				& "'," _
	      & "'" & ""						& "'," _
	      &       locTax					& " ," _
	      &       locShipZone				& " ," _
	      & "'" & locStatus					& "' " _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	'Get idProduct of INSERTed Record
	mySQL = "SELECT MAX(idLocation) AS maxIdLocation " _
		  & "FROM   Locations "
	set rs = openRSexecute(mySQL)
	idLocation = rs("maxIdLocation")
	call closeRS(rs)
	
	call closedb()
	Response.Redirect "SA_loc_edit.asp?action=edit&recID=" & idLocation & "&msg=" & server.URLEncode("Country was added. You may Add or Edit States/Provinces if required.")
	
end if

'DELETE - Delete Country and/or State(s)
if action = "del" then

	'Delete Country and all related State record(s)
	if len(locState) = 0 or IsNull(locState) then	
		mySQL = "DELETE FROM Locations " _
		      & "WHERE locCountry = '" & locCountry & "' "
		set rs = openRSexecute(mySQL)
		
		call closedb()
		Response.Redirect "SA_loc.asp?recallCookie=1&msg=" & server.URLEncode("Country was deleted.")
		
	'Delete only State record
	else						
		mySQL = "DELETE FROM Locations " _
		      & "WHERE locCountry = '" & locCountry & "' " _
		      & "AND   locState   = '" & locState   & "' "
		set rs = openRSexecute(mySQL)
		
		'Get idLocation of Country record so we can redirect
		mySQL = "SELECT idLocation " _
			  & "FROM   Locations " _
			  & "WHERE  locCountry = '" & locCountry & "' " _
			  & "AND   (locState IS NULL OR locState = '') "
		set rs = openRSexecute(mySQL)
		idLocation = rs("idLocation")
		call closeRS(rs)
		
		call closedb()
		Response.Redirect "SA_loc_edit.asp?action=edit&recID=" & idLocation & "&msg=" & server.URLEncode("State/Province was deleted.")

	end if
	
end if

'BULK DELETE - Delete Multiple Countries and related State(s)
if action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of locCountries that will be deleted

	'Create array of countries to be deleted
	delArray = split(Request.Form("locCountry"),",")
	
	'Loop through list of countries and delete one by one
	for delI = LBound(delArray) to UBound(delArray)

		mySQL = "DELETE FROM Locations " _
		      & "WHERE locCountry = '" & trim(delArray(delI)) & "' "
		set rs = openRSexecute(mySQL)
	
	next
		
	call closedb()
	Response.Redirect "SA_loc.asp?recallCookie=1&msg=" & server.URLEncode("Countries were deleted.")
	
end if

'EDIT - Edit Country
if action = "edit" then

	'Get existing (old) Country Code of edited record
	mySQL = "SELECT locCountry " _
		  & "FROM   Locations " _
		  & "WHERE  idLocation = " & idLocation & " "
	set rs = openRSexecute(mySQL)
	locCountryOld = rs("locCountry")
	call closeRS(rs)

	'Set CursorLocation of the Connection Object to Client
	cn.CursorLocation = adUseClient
	
	'BEGIN Transaction
	cn.BeginTrans
	
	'Update Country record
	mySQL = "UPDATE Locations SET " _
	      & "locCountry='"    & locCountry					& "'," _
	      & "locState='"      &	""							& "'," _
	      & "locName='"		  & replace(locName,"'","''")	& "'," _
	      & "locTax="		  & locTax						& " ," _
	      & "locShipZone="	  & locShipZone					& " ," _
	      & "locStatus='"	  & locStatus					& "' " _
	      & "WHERE idLocation = " & idLocation
	set rs = openRSexecute(mySQL)
	
	'Update State/Province records with new Country Code
	mySQL = "UPDATE Locations SET " _
	      & "locCountry = '" & locCountry & "' " _
	      & "WHERE locCountry = '" & locCountryOld & "' "
	set rs = openRSexecute(mySQL)

	'END Transaction
	cn.CommitTrans

	call closedb()
	Response.Redirect "SA_loc_edit.asp?action=edit&recID=" & idLocation & "&msg=" & server.URLEncode("Country was updated.")

end if

'ADDSTATE - Add State
if action = "addstate" then

	'Add State/Province 
	mySQL = "INSERT INTO Locations (" _
		  & "locName,locCountry,locState,locTax,locShipZone,locStatus" _
	      & ") VALUES (" _
	      & "'" & replace(locName,"'","''") & "'," _
	      & "'" & locCountry				& "'," _
	      & "'" & locState					& "'," _
	      &       locTax					& " ," _
	      &       locShipZone				& " ," _
	      & "'" & locStatus					& "' " _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	'Get idLocation of Country record so we can redirect
	mySQL = "SELECT idLocation " _
		  & "FROM   Locations " _
		  & "WHERE  locCountry = '" & locCountry & "' " _
		  & "AND   (locState IS NULL OR locState = '') "
	set rs = openRSexecute(mySQL)
	idLocation = rs("idLocation")
	call closeRS(rs)
	
	call closedb()
	Response.Redirect "SA_loc_edit.asp?action=edit&recID=" & idLocation & "&msg=" & server.URLEncode("State/Province was added.")

end if

'EDITSTATE - Edit State
if action = "editstate" then

	'Update State/Province record
	mySQL = "UPDATE Locations SET " _
	      & "locState='"      &	locState					& "'," _
	      & "locName='"		  & replace(locName,"'","''")	& "'," _
	      & "locTax="		  & locTax						& " ," _
	      & "locShipZone="	  & locShipZone					& " ," _
	      & "locStatus='"	  & locStatus					& "' " _
	      & "WHERE idLocation = " & idLocation
	set rs = openRSexecute(mySQL)
	
	'Get idLocation of Country record so we can redirect
	mySQL = "SELECT idLocation " _
		  & "FROM   Locations " _
		  & "WHERE  locCountry = '" & locCountry & "' " _
		  & "AND   (locState IS NULL OR locState = '') "
	set rs = openRSexecute(mySQL)
	idLocation = rs("idLocation")
	call closeRS(rs)
	
	call closedb()
	Response.Redirect "SA_loc_edit.asp?action=edit&recID=" & idLocation & "&msg=" & server.URLEncode("State/Province was edited.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_loc.asp?recallCookie=1"

%>
