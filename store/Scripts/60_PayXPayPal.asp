<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : PayPal IPN handler. Updates order as required.
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
'1. We don't call the sessionCart() and sessionCust() functions on 
'   this page because we don't know when the page will be called. 
'   PayPal can call this page at any time depending on when 
'   their system updates the transaction. Besides, we get the order 
'   number from PayPal so these functions are not really needed.
'2. This script is "silent", in other words no HTML output is 
'   actually sent to the user's browser when called by PayPal.
'*************************************************************************
Option explicit
Response.Buffer = true
%>
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="_INCconfig_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="_INCappFunctions_.asp"-->
<!--#include file="_INCappEmail_.asp"-->
<!--#include file="_INCupdStatus_.asp"-->
<%
'Work Fields
dim strReply
dim objHttp
dim httpStatus
dim httpResponseText
dim formattedDateTime

'Values sent by PayPal
dim payment_status
dim receiver_email
dim item_number

'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2

'Session
dim idOrder
dim idCust

'************************************************************************
 
'Open Database Connection
call openDb()

'Store Configuration
if loadConfig() = false then
	call errorDB(langErrConfig,"")
end if

'Get a Date and Time formatted to the user's specifications
formattedDateTime = formatTheDate(now()) & " " & time()

'Read post from PayPal system and add 'cmd'
strReply = Request.Form & "&cmd=_notify-validate"

'Extract some info we will need to update the Database if "VERIFIED"
payment_status	= trim(Request.Form("payment_status"))
receiver_email	= trim(Request.Form("receiver_email"))
item_number		= trim(Request.Form("item_number"))

'Check the item_number and Order Status
if len(item_number) = 0 or not isNumeric(item_number) then
	Response.Clear
	Response.Write "PayPal IPN : Invalid value in Item_Number field."
	Response.End
else
	'Check that the order status is "Pending"
	mySQL = "SELECT orderStatus " _
		  & "FROM   cartHead " _
		  & "WHERE  idOrder = " & validSQL(item_number,"I")
	set rsTemp = openRSexecute(mySQL)
	if rsTemp.eof then
		Response.Clear
		Response.Write "PayPal IPN : Order could not be located."
		Response.End
	else
		if rsTemp("orderStatus") <> "0" then 
			Response.Clear
			Response.Write "PayPal IPN : Order status must be pending."
			Response.End
		end if
	end if
	call closeRS(rsTemp)
end if

'Create XML object
on error resume next
set objHttp = server.Createobject("MSXML2.ServerXMLHTTP")
if err.number <> 0 then
	err.Clear
	set objHttp = server.Createobject("MSXML2.ServerXMLHTTP.4.0")
	if err.number <> 0 then
		Response.Clear
		Response.Write "PayPal IPN : Could not create XML HTTP object."
		Response.End
	end if
end if
on error goto 0

'Open connection to LIVE server (PayPal)
objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false

'Open connection to TEST server (EliteWeaver)
'objHttp.open "POST", "http://www.eliteweaver.co.uk/testing/ipntest.php", false
'objHttp.setRequestHeader "Host","www.eliteweaver.co.uk"
'objHttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"

'Send reply
objHttp.Send strReply

'Get response
httpStatus		 = objHttp.status
httpResponseText = UCase(trim(objHttp.responseText))
set objHttp		 = nothing

'Validate response
if httpStatus <> 200 then
	call updOrderPrivate(item_number,"DATE : " & formattedDateTime & vbCrLf & "PayPal IPN : HTTP Error " & httpStatus)
	Response.Write "PayPal IPN : HTTP Error " & httpStatus
else
	if httpResponseText = "VERIFIED" then
		if lCase(payPalMemberID) = lCase(receiver_email) then
			if lCase(payment_status) = "completed" then
				call updOrderStatus(item_number,"1","Y","Y","DATE : " & formattedDateTime & vbCrLf & "PayPal IPN : Status = " & payment_status)
				Response.Write "PayPal IPN : Status = " & payment_status
			else
				call updOrderPrivate(item_number,"DATE : " & formattedDateTime & vbCrLf & "PayPal IPN : Status = " & payment_status)
				Response.Write "PayPal IPN : Status = " & payment_status
			end if
		else
			call updOrderPrivate(item_number,"DATE : " & formattedDateTime & vbCrLf & "PayPal IPN : Invalid Email = " & receiver_email)
			Response.Write "PayPal IPN : Invalid Email = " & receiver_email
		end if
	else
		if httpResponseText = "INVALID" then
			call updOrderPrivate(item_number,"DATE : " & formattedDateTime & vbCrLf & "PayPal IPN : INVALID Response")
			Response.Write "PayPal IPN : INVALID Response"
		else
			call updOrderPrivate(item_number,"DATE : " & formattedDateTime & vbCrLf & "PayPal IPN : ERROR Response")
			Response.Write "PayPal IPN : ERROR Response"
		end if
	end if
end if

'Close Database connection
call closeDB()
%>