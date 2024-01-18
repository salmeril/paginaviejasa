<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Gateway between HTTP and HTTPS sessions
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
'Variables
dim action
dim randomKey

'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2

'Session
dim idOrder
dim idCust

'*************************************************************************

'Open Database Connection
call openDb()

'Store Configuration
if loadConfig() = false then
	call errorDB(langErrConfig,"")
end if

'Check Action Indicator
action = lCase(Request.QueryString("action"))
if  action <> "logon"    _
and action <> "retrieve" _
and action <> "save"     _
and action <> "checkout" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrAction)
end if

'********************************
'* HTTP -> HTTPS (10_logon.asp) *
'********************************
if action = "checkout" _
or action = "save"     _
or action = "logon"    then

	'Get idOrder from Session
	idOrder = sessionCart()
	
	'Check if there is an active Shopping Cart
	if isNull(idOrder) then

		'Close DB Connection
		call closedb()

		if action = "logon" then
			'Redirect to "10_logon.asp" without passing session
			Response.Redirect urlSSL & "10_Logon.asp?action=" & server.URLEncode(action)
		else
			'Show error message
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrCartEmpty)
		end if

	end if

	'Generate a Random Key. This key is passed along with the order 
	'number to ensure that the Order Number is not tampered with.
	randomKey = rndKey(99999999)

	'Update cartHead with key so we can check it in "10_logon.asp"
	mySQL = "UPDATE cartHead " _
	      & "SET    randomKey = '" & validSQL(randomKey,"A") & "' " _
		  & "WHERE  idOrder = "    & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)

	'Close DB Connection
	call closedb()

	'Redirect to "10_logon.asp" and pass session
	Response.Redirect urlSSL & "10_Logon.asp?action=" & server.URLEncode(action) & "&idOrder=" & idOrder & "&randomKey=" & randomKey
		
'****************************
'* HTTPS -> HTTP (cart.asp) *
'****************************
else
	
	'Validate Order Number passed via QueryString
	idOrder = Request.QueryString("idOrder")
	if not isNumeric(idOrder) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
	end if

	'Validate Random Key passed via QueryString
	randomKey = Request.QueryString("randomKey")
	if not isNumeric(randomKey) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvRandKey)
	end if

	'Validate Order/Random Key/Status combination on DB
	mySQL = "SELECT idOrder " _
		  & "FROM   cartHead " _
		  & "WHERE  idOrder = "    & validSQL(idOrder,"I")   & " " _
		  & "AND    randomKey = '" & validSQL(randomKey,"A") & "' " _
		  & "AND   (orderStatus = 'U' OR orderStatus = 'S') "
	set rsTemp = openRSexecute(mySQL)
	if rstemp.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
	end if
	call closeRS(rsTemp)

	'Set Session Value
	session(storeID & "idOrder") = idOrder

	'Close DB Connection
	call closedb()

	'Redirect to "cart.asp"
	Response.Redirect urlNonSSL & "cart.asp"
	
end if

'**********************************************************************
'Generate a Random Key
'**********************************************************************
function rndKey(upperbound)
	randomize
	rndKey = DatePart("y",now()) _
		   & DatePart("h",now()) _
		   & DatePart("n",now()) _
		   & DatePart("s",now()) _
		   & Int(upperbound * Rnd + 1)
end function

%>