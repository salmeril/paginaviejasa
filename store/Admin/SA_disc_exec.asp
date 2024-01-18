<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Order Discount Maintenance
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
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="../Scripts/_INCappFunctions_.asp"-->
<%
'Database
dim mySQL, cn, rs

'DiscOrder
dim idDiscOrder
dim discCode
dim discPerc
dim discAmt
dim discFromAmt
dim discToAmt
dim discStatus
dim discOnceOnly
dim discValidFrom
dim discValidTo

'Work Fields
dim action
dim discValidFromDD
dim discValidFromMM
dim discValidFromYYYY
dim discValidToDD
dim discValidToMM
dim discValidToYYYY

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
if action <> "edit" and action <> "add" and action <> "bulkdel" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idDiscOrder
if action = "edit" then

	idDiscOrder = trim(Request.Form("idDiscOrder"))
	if len(idDiscOrder) = 0 then
		idDiscOrder = trim(Request.QueryString("idDiscOrder"))
	end if
	if idDiscOrder = "" or not isNumeric(idDiscOrder) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Order Discount ID.")
	else
		idDiscOrder = CLng(idDiscOrder)
	end if
	
end if

if action = "edit" or action = "add" then

	'Get discCode
	discCode = trim(Request.Form("discCode"))
	if len(discCode) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Discount Code.")
	end if
	if inStr(discCode," ")  > 0 _
	or inStr(discCode,"'")  > 0 _
	or inStr(discCode,"""") > 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid characters in the Discount Code.")
	end if
	
	'Check discCode not a duplicate
	mySQL = "SELECT discCode " _
	      & "FROM   DiscOrder " _
	      & "WHERE  discCode = '" & discCode & "' "
	if action = "edit" then
		mySQL = mySQL & "AND idDiscOrder <> " & idDiscOrder
	end if
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Discount Code already exists.")
	end if
	call closeRS(rs)
	
	'Get discFromAmt
	discFromAmt = trim(Request.Form("discFromAmt"))
	if len(discFromAmt) = 0 or not Isnumeric(discFromAmt) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Order Amount FROM value.")
	end if
	discFromAmt = CDbl(discFromAmt)
	if discFromAmt < 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Order Amount FROM value.")
	end if
	
	'Get discToAmt
	discToAmt = trim(Request.Form("discToAmt"))
	if len(discToAmt) = 0 or not Isnumeric(discToAmt) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Order Amount TO value.")
	end if
	discToAmt = CDbl(discToAmt)
	if discToAmt < 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Order Amount TO value.")
	end if
	
	'Check TO is greater than FROM
	if discToAmt < discFromAmt then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Order Amount TO value must be greater that Order Amount FROM value.")
	end if
	
	'Get discPerc and/or discAmt
	discPerc = trim(Request.Form("discPerc"))
	discAmt  = trim(Request.Form("discAmt"))
	if (len(discPerc) = 0 and len(discAmt) = 0) _
	or (len(discPerc) > 0 and len(discAmt) > 0) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Discount Percentage and/or Discount Amount.")
	'Discount Percentage
	elseif len(discPerc) > 0 then
		if not Isnumeric(discPerc) then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Discount Percentage.")
		end if
		discPerc = CDbl(discPerc)
		discAmt  = null
		if discPerc <= 0 or discPerc > 100 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Discount Percentage.")
		end if
	'Discount Amount
	else
		if not Isnumeric(discAmt) then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Discount Amount.")
		end if
		discAmt	 = CDbl(discAmt)
		discPerc = null
		if discAmt <= 0 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Discount Amount.")
		end if
		if discAmt > discToAmt then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Discount Amount can not be greater than To amount.")
		end if
	end if
	
	'Get discStatus
	discStatus = UCase(trim(Request.Form("discStatus")))
	if discStatus <> "A" and discStatus <> "I" and discStatus <> "U" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Status.")
	end if
	
	'Get discOnceOnly
	discOnceOnly = UCase(trim(Request.Form("discOnceOnly")))
	if discOnceOnly <> "Y" and discOnceOnly <> "N" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Once Only value.")
	end if
	
	'Get discValidFrom
	discValidFromDD   = trim(Request.Form("discValidFromDD"))
	discValidFromMM   = trim(Request.Form("discValidFromMM"))
	discValidFromYYYY = trim(Request.Form("discValidFromYYYY"))
	if not (len(discValidFromDD)=2 and len(discValidFromMM)=2 and len(discValidFromYYYY)=4) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid FROM Date.")
	end if
	if not isDate(discValidFromMM & "/" & discValidFromDD & "/" & discValidFromYYYY) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid FROM Date.")
	end if
	discValidFrom = discValidFromYYYY & discValidFromMM & discValidFromDD

	'Get discValidTo
	discValidToDD   = trim(Request.Form("discValidToDD"))
	discValidToMM   = trim(Request.Form("discValidToMM"))
	discValidToYYYY = trim(Request.Form("discValidToYYYY"))
	if not (len(discValidToDD)=2 and len(discValidToMM)=2 and len(discValidToYYYY)=4) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid TO Date.")
	end if
	if not isDate(discValidToMM & "/" & discValidToDD & "/" & discValidToYYYY) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid TO Date.")
	end if
	discValidTo = discValidToYYYY & discValidToMM & discValidToDD
	
	'Check TO is greater than FROM. Note : Dates are stored as 
	'integer strings, hence the use of CLng() below
	if CLng(discValidTo) < CLng(discValidFrom) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("TO Date can not be less than FROM Date.")
	end if

end if

'ADD
if action = "add" then

	'Add Order Discount
	mySQL = "INSERT INTO DiscOrder (" _
		  & "discCode,discPerc,discAmt,discFromAmt," _
		  & "discToAmt,discStatus,discOnceOnly," _
		  & "discValidFrom,discValidTo" _
	      & ") VALUES (" _
	      & "'" & discCode						& "'," _
	      &       emptyString(discPerc,"NULL")	& "," _
	      &       emptyString(discAmt,"NULL")	& "," _
	      &       discFromAmt					& "," _
	      &       discToAmt						& "," _
	      & "'" & discStatus					& "'," _
	      & "'" & discOnceOnly					& "'," _
	      & "'" & discValidFrom					& "'," _
	      & "'" & discValidTo					& "'" _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	call closedb()
	Response.Redirect "SA_disc.asp?recallCookie=1&msg=" & server.URLEncode("Order Discount record was Added.")
	
end if

'BULK DELETE
if action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of idOrders that will be deleted

	'Populate array with list of records selected for deletion.
	delArray = split(Request.Form("idDiscOrder"),",")
	
	'Loop through list of orders and delete one by one
	for delI = LBound(delArray) to UBound(delArray)

		'Delete record from shipRates
		mySQL = "DELETE FROM DiscOrder " _
		      & "WHERE idDiscOrder = " & trim(delArray(delI))
		set rs = openRSexecute(mySQL)
		
	next

	call closedb()
	Response.Redirect "SA_disc.asp?recallCookie=1&msg=" & server.URLEncode("Order Discount record(s) were Deleted.")

end if

'EDIT
if action = "edit" then

	'Update Order Discount
	mySQL = "UPDATE DiscOrder SET " _
	      & "discCode='"		& discCode						& "', " _
	      & "discPerc="			& emptyString(discPerc,"NULL")	& ",  " _
	      & "discAmt="			& emptyString(discAmt,"NULL")	& ",  " _
	      & "discFromAmt="		& discFromAmt					& ",  " _
	      & "discToAmt="		& discToAmt						& ",  " _
	      & "discStatus='"		& discStatus					& "', " _
	      & "discOnceOnly='"	& discOnceOnly					& "', " _
	      & "discValidFrom='"	& discValidFrom					& "', " _
	      & "discValidTo='"		& discValidTo					& "'  " _
	      & "WHERE idDiscOrder = " & idDiscOrder
	set rs = openRSexecute(mySQL)
	
	call closedb()
	Response.Redirect "SA_disc.asp?recallCookie=1&msg=" & server.URLEncode("Order Discount record was Updated.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_disc.asp?recallCookie=1"

%>
