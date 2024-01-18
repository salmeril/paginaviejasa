<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Logon/Logoff existing customers
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
<!--#include file="_INCrc4_.asp"-->
<%
'Work fields
dim strReferer
dim arrayErrors
dim action
dim randomKey
dim formID

'Customer
dim status
dim Name
dim LastName
dim Email
dim Password

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

'Get action indicator
action = lCase(Request.QueryString("action"))
if len(action) = 0 then
	action = lCase(Request.Form("action"))
end if

'LogOff?
if action = "logoff" then
	'Log the user off, and give a message
	session(storeID & "idCust")	= null
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langGenLogoffMsg)
	
'Logon, Save, or Checkout?
else
	if  action <> "logon" _
	and action <> "save" _
	and action <> "checkout" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrAction)
	end if
end if

'***********************************************************************
'This script will typically run under an HTTPS session. Because 
'sessions are not shared between HTTP and HTTPS, we execute the code 
'below to re-create the session under HTTPS if needed. To make sure 
'that the Order Number was not altered by the user, we compare the 
'Random Key in the database (generated in 05_Gateway.asp), against 
'the Random Key passed via the querystring.
'***********************************************************************
idOrder   = trim(Request.QueryString("idOrder"))
randomKey = trim(Request.QueryString("randomKey"))
if len(idOrder) > 0 and len(randomKey) > 0 then

	'Validate Order Number passed via QueryString
	if not isNumeric(idOrder) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
	end if

	'Validate Random Key passed via QueryString
	if not isNumeric(randomKey) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvRandKey)
	end if

	'Validate Order/Random Key combination on DB
	mySQL = "SELECT idOrder " _
		  & "FROM   cartHead " _
		  & "WHERE  idOrder = "    & validSQL(idOrder,"I")   & " " _
		  & "AND    randomKey = '" & validSQL(randomKey,"A") & "' "
	set rsTemp = openRSexecute(mySQL)
	if rstemp.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
	end if
	call closeRS(rsTemp)

	'Create session
	session(storeID & "idOrder") = idOrder
	
end if

'Get/Set Cart/Order Session
idOrder = sessionCart()

'Get/Set Customer Session
idCust  = sessionCust()

'If Checkout or Save, do some validations.
if action = "checkout" or action = "save" then

	'Check if the session is still active
	if isNull(idOrder) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrCartEmpty)
	end if

	'Check if cart has any items
	if cartQty(idOrder) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrCartEmpty)
	end if

	'Check if minimum order amount has been met (checkout only)
	if action = "checkout" then
		if cartTotal(idOrder,0) < pMinCartAmount then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrMinPrice & pCurrencySign & moneyS(pMinCartAmount))
		end if
	end if
	
end if

'Get Form ID
formID = trim(Request.Form("formID"))

'Before we display the form for the first time, do some checks
if formID = "" then

	'If user is already logged on, forward to next page
	if not(isNull(idCust)) then
		select case action
		case "logon"
			Response.Redirect "custListOrders.asp"
		case "save"
			call saveCart(idOrder,idCust)
			Response.Redirect "custListOrders.asp"
		case "checkout"
			Response.Redirect "20_Customer.asp?action=checkout"
		end select
	end if
	
end if

'Check if the Customer clicked the "New Customer" button
if formID = "00a" then

	'Forward to next page
	select case action
	case "logon"
		Response.Redirect "20_Customer.asp?action=newacc"
	case "save"
		Response.Redirect "20_Customer.asp?action=save"
	case "checkout"
		Response.Redirect "20_Customer.asp?action=checkout"
	end select

end if

'Check if the Customer clicked the "Logon" button
if formID = "00" then

	'Get values from the form
	Email	 = request.form("Email")
	Password = request.form("Password")

	'Check Customer Logon form
	if len(email) = 0 or len(password) = 0 then
		'Email OR Password is empty
		if len(email) = 0 then
			arrayErrors = arrayErrors & "|email"
		end if
		if len(password) = 0 then
			arrayErrors = arrayErrors & "|password"
		end if
	else
		'Check for Invalid Characters
		if invalidChar(Email,1,"@.-_") then
			arrayErrors = arrayErrors & "|email"
		end if
		if invalidChar(Password,1,"") then
			arrayErrors = arrayErrors & "|password"
		end if
	end if
	
	'If there weren't any obvious errors so far, check the Customer's 
	'DB record and if it's valid log him on.
	if len(trim(arrayErrors)) = 0 then
	
		'Check Email/Password, and if Customer is still Active
		mySQL = "SELECT idCust " _
			  & "FROM   customer " _
			  & "WHERE  email='"    & validSQL(email,"A") & "' " _
			  & "AND    password='" & validSQL(Ascii2Hex(EnDeCrypt(lCase(password),rc4Key)),"A") & "' " _
			  & "AND    status='A'"
		set rsTemp = openRSexecute(mySQL)
		if not rstemp.eof then
		
			'Log the Customer on
			idCust = rsTemp("idCust")
			session(storeID & "idCust") = idCust
			
			'Forward to next page
			select case action
			case "logon"
				Response.Redirect "custListOrders.asp"
			case "save"
				call saveCart(idOrder,idCust)
				Response.Redirect "custListOrders.asp"
			case "checkout"
				Response.Redirect "20_Customer.asp?action=checkout"
			end select

		else
		
			'Invalid email/password combo entered
			arrayErrors = arrayErrors & "|email"
			arrayErrors = arrayErrors & "|password"
			
		end if
		call closeRS(rsTemp)

	end if
	
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<!-- Outer Table Cell -->
<table border="0" cellpadding="0" cellspacing="0" width=350><tr><td>

<!-- Inner Table -->
<table border="0" cellspacing="0" cellpadding="2" width="100%">
	<tr>
		<td valign=middle class="CPpageHead" width="49%">
			<b><%=langGenExistCust%></b>
		</td>
		<td width="2%">&nbsp;</td>
		<td valign=middle class="CPpageHead" width="49%">
			<b><%=langGenNewCust%></b>
		</td>
	</tr>
	<tr>
		<form METHOD="POST" name="LogonForm" action="10_Logon.asp">
		<td valign=top>
			<br>
<%
			'If there were errors, show message
			if len(trim(arrayErrors)) > 0 then
				arrayErrors = split(LCase(arrayErrors),"|")
				Response.Write "<font color=red><i>" & langErrInvForm & "</i></font><br><br>"
			else
				arrayErrors = array("")
			end if
%>
			<%=langGenEmail & " " & checkFieldError("email",arrayErrors)%><br>
			<input type=text name=email size=20 maxlength=50 value="<%=email%>"><br>
			<%=langGenPassword & " " & checkFieldError("password",arrayErrors)%><br>
			<input type=password name=password size=10 maxlength=10 value="<%=password%>"><br><br>
			<input type="hidden" name="action" value="<%=action%>">
			<input type="hidden" name="formID" value="00">
			<input type="submit" name="Submit" value="<%=langGenLogon%>"><br><br>
			<a href="<%=urlNonSSL%>custPass.asp"><i><%=langGenForgetPass%></i></a>
		</td>
		</form>
		
		<td>&nbsp;</td>
		
		<form METHOD="POST" name="NewAccForm" action="10_Logon.asp">
		<td valign=top>
			<br>
			<%=langGenNewCustDesc%><br><br>
			<input type="hidden" name="action" value="<%=action%>">
			<input type="hidden" name="formID" value="00a">
			<input type="submit" name="Submit" value="<%=langGenNewCust%>">
		</td>
		</form>
	</tr>
</table>

<!-- End Outer Table Cell -->
</td></tr></table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%
call closedb()
%>