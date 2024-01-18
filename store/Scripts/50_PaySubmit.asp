<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Request payment from 3rd party payment processors
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
<!--#include file="_INCauthNet_.asp"-->
<%
'cartHead
dim orderStatus
dim orderDate
dim subTotal
dim taxTotal
dim shipmentTotal
dim handlingFeeTotal
dim Total
dim Name
dim LastName
dim CustomerCompany
dim Phone
dim Email
dim Address
dim City
dim Zip
dim locState
dim locCountry
dim cardNumber
dim paymentType

'Work Fields
dim countryCode
dim stateCode
dim f
dim qIdOrder
dim refererURL

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

'Get/Set Cart/Order Session
idOrder = sessionCart()

'Get/Set Customer Session
idCust = sessionCust()

'Check that the Customer is currently logged in
if isNull(idCust) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrNotLoggedIn)
end if

'NOTE : Some gateways require that this page's URL be fixed (eg. 
'LinkPoint). We can therefore NOT pass any variable info to this 
'script in a querystring. This information must be passed via the 
'session object, or via a POST action from a form.

'Get Order Number and Validate
qIdOrder = session(storeID & "idOrderPaySubmit")
if len(qIdOrder) = 0 then
	qIdOrder = Request.Form("idOrder")
end if
if len(qIdOrder) = 0 then
	qIdOrder = Request.QueryString("idOrder")
end if
if len(qIdOrder) = 0 or not IsNumeric(qIdOrder) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
end if

'Retrieve some information we may need from cartHead
mySQL="SELECT orderStatus,orderDate,subTotal," _
	& "       taxTotal,shipmentTotal,Total," _
	& "       Name,LastName,CustomerCompany," _
	& "       Phone,Email,Address," _
	& "       City,Zip,locState," _
	& "       locCountry,cardNumber,paymentType," _
	& "       handlingFeeTotal " _
	& "FROM   cartHead " _
	& "WHERE  idOrder = " & validSQL(qIdOrder,"I") & " " _
	& "AND    idCust = "  & validSQL(idCust,"I")
set rsTemp = openRSexecute(mySQL)
if not rstemp.eof then
	orderStatus			= rstemp("orderStatus")
	orderDate			= rstemp("orderDate")
	subTotal			= rstemp("subTotal")
	taxTotal			= rstemp("taxTotal")
	shipmentTotal		= rstemp("shipmentTotal")
	Total				= rstemp("Total")
	Name				= trim(rstemp("name"))
	LastName			= trim(rstemp("LastName"))
	CustomerCompany		= trim(rstemp("CustomerCompany"))
	Phone				= trim(rstemp("Phone"))
	Email				= trim(rstemp("Email"))
	Address				= trim(rstemp("Address"))
	City				= trim(rstemp("City"))
	Zip					= trim(rstemp("Zip"))
	locState			= trim(rstemp("locState"))
	locCountry			= trim(rstemp("locCountry"))
	cardNumber			= trim(rstemp("cardNumber"))
	paymentType			= trim(rstemp("paymentType"))
	handlingFeeTotal	= rstemp("handlingFeeTotal")
else
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
end if
call closeRS(rsTemp)

'The order will have the full country and state description. Retrieve 
'the country and state codes for the benefit of some payment processors 
'that require the country and state codes, and not the description.
countryCode = getCountryCode(locCountry)
stateCode   = getStateCode(locState,countryCode)

'Validate Payment Processor(s)
if  lCase(paymentType) <> "paypal" _
and lCase(paymentType) <> "2checkout" _
and lCase(paymentType) <> "authorizenet" _
and lCase(paymentType) <> "custom" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvPayment)
end if

'Validate Order Status
if orderStatus <> "0" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrdStat)
end if

'What page did we come from?
refererURL = lCase(Request.ServerVariables("HTTP_REFERER"))

%> 
<!--#include file="../UserMods/_INCtop_.asp"-->

<!-- Outer Table Cell -->
<table border="0" cellpadding="0" cellspacing="0" width=450><tr><td>

<!-- Heading -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
<tr><td nowrap valign=middle class="CPpageHead">
	<table border=0 cellpadding=0 cellspacing=0 width="100%">
	<tr>
		<td nowrap valign=middle>
			<b><%=langGenSubmitPayment%></b>
		</td>
		<td nowrap align=right valign=middle>
<%
			'Determine if this is part of the checkout process
			if instr(refererURL,lCase("40_SubmitOrder.asp")) <> 0 then
%>
				<b><font color=#800000>[ <%=langGenStep%> 4 / 4 ]</font></b>
<%
			else
				Response.Write "&nbsp;"
			end if
%>
		</td>
	</tr>
	</table>
</td></tr>
</table>
	
<img src="../UserMods/misc_cleardot.gif" height=4 width=1><br>

<!-- Payment Button -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
<%
	'What Payment Processor?
	if lCase(paymentType) = "paypal" then
		call payPayPal()
	end if
	if lCase(paymentType) = "2checkout" then
		call pay2CheckOut()
	end if
	if lCase(paymentType) = "authorizenet" then
		call payAuthorizeNet()
	end if
	if lCase(paymentType) = "custom" then
		call payCustomRoutine()
	end if
%>
</table>

<br>

<!-- Payment Summary -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr><td class="CPpageHead" valign=middle>	
		<b><%=langGenOrderSummary%></b>
	</td></tr>
</table>

<table border="0" cellpadding="2" cellspacing="0">
	<tr> 
		<td valign=top nowrap><b><%=langGenFullName%></b>&nbsp;</td>
		<td valign=top><%=name & " " & lastname%></td>
	</tr>
	<tr>
		<td valign=top nowrap><b><%=langGenOrderNumber%></b>&nbsp;</td>
		<td valign=top><%=pOrderPrefix & "-" & qIdOrder%></td>
	</tr>
	<tr>
		<td valign=top nowrap><b><%=langGenOrderDate%></b>&nbsp;</td>
		<td valign=top><%=formatTheDate(orderDate)%></td>
	</tr>
	<tr>
		<td valign=top nowrap><b><%=langGenTotal%></b>&nbsp;</td>
		<td valign=top><%=pCurrencySign & moneyS(Total)%></td>
	</tr>
	<tr>
		<td valign=top nowrap><b><%=langGenPayment%></b>&nbsp;</td>
		<td valign=top><%=paymentMsg(paymentType, total, cardNumber)%></td>
	</tr>
</table>
   
<!-- End Outer Table Cell -->
</td></tr></table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%

call closeDB()
'**********************************************************************
'PayPal payments
'Notes : Relies on you specifying seperate URL's for transactions that 
'        are successful and unsuccessful. Because we have one script 
'        that deals with both, we construct the two return URL's to 
'        go to the same page, but we add a status indicator to the 
'        URL.
'**********************************************************************
sub payPayPal()

	'Check if we are in demo mode
	if demoMode = "Y" then
%>
	<form method="post" action="<%=urlSsl%>60_PayReturn.asp?CP_idOrder=<%=qIDOrder%>&CP_Status=success" id=PayPalForm name=PayPalForm>
<%
	else
%>
	<form method="post" action="https://www.paypal.com/cgi-bin/webscr" id=PayPalForm name=PayPalForm>
<%
	end if
%>
	<TR> 
		<TD class="CPgenHeadings">
			<input type="hidden" name="cmd"					value="_ext-enter">
			<input type="hidden" name="redirect_cmd"		value="_xclick">
			<input type="hidden" name="business"			value="<%=payPalMemberID%>">
			<input type="hidden" name="item_name"			value="<%=pCompany & " Order " & pOrderPrefix & "-" & qIdOrder%>">
			<input type="hidden" name="item_number"			value="<%=qIdOrder%>">
			<input type="hidden" name="amount"				value="<%=moneyD(total)%>">
			<input type="hidden" name="currency_code"		value="<%=payPalCurrCode%>">
			<input type="hidden" name="first_name"			value="<%=name%>">
			<input type="hidden" name="last_name"			value="<%=lastName%>">
			<input type="hidden" name="address1"			value="<%=address%>">
			<input type="hidden" name="zip"					value="<%=zip%>">
			<input type="hidden" name="city"				value="<%=city%>">
			<input type="hidden" name="email"				value="<%=email%>">
			<input type="hidden" name="return"				value="<%=urlSsl%>60_PayReturn.asp?CP_idOrder=<%=qIDOrder%>&CP_Status=success">
			<input type="hidden" name="cancel_return"		value="<%=urlSsl%>60_PayReturn.asp?CP_idOrder=<%=qIDOrder%>&CP_Status=error">
			<input type="hidden" name="no_shipping"			value="1">
			<input type="hidden" name="no_note"				value="1">
			<input type="hidden" name="undefined_quantity"	value="0">
			<br>
			<center>
				<b><font color=red><%=langGenPayNowMsg%></font></b>
				<br><br>
				<b><font color=red size=2>--&gt;&nbsp;&nbsp;&nbsp;</font></b>
				<input type="image" src="../UserMods/butt_PayPal.gif" border="0" name="submit" alt="Submit Payment" align="middle">
				<b><font color=red size=2>&nbsp;&nbsp;&nbsp;&lt;--</font></b>
			</center>
			<br>
		</TD>
	</TR>
	</form>
<%
end sub
'**********************************************************************
'2CheckOut payments
'Notes : Always returns control to the same URL, regardless of the 
'        status of the transaction. The return URL has to be entered 
'        into your 2CheckOut account settings. When control is 
'        returned,  2CheckOut passes a status indicator which can be 
'        checked.
'**********************************************************************
sub pay2CheckOut()

	'Check if we are in demo mode
	if demoMode = "Y" then
%>
	<form method="post" action="<%=urlSsl%>60_PayReturn.asp?CP_idOrder=<%=qIDOrder%>&CP_Status=success" id=TwoCheckOutForm name=TwoCheckOutForm>
<%
	else
%>
	<form method="post" action="https://www.2checkout.com/cgi-bin/sbuyers/cartpurchase.2c" id=TwoCheckOutForm name=TwoCheckOutForm>
<%
	end if
%>
	<TR> 
		<TD class="CPgenHeadings">
			<input type="hidden" name="sid"					value="<%=TwoCheckOutSID%>">
			<input type="hidden" name="total"				value="<%=moneyD(total)%>">
			<input type="hidden" name="cart_order_id"		value="<%=qIdOrder%>">
			<input type="hidden" name="card_holder_name"	value="<%=name & " " & lastName%>">
			<input type="hidden" name="street_address"		value="<%=address%>">
			<input type="hidden" name="city"				value="<%=city%>">
			<input type="hidden" name="state"				value="<%=locState%>">
			<input type="hidden" name="zip"					value="<%=zip%>">
			<input type="hidden" name="country"				value="<%=locCountry%>">
			<input type="hidden" name="email"				value="<%=email%>">
			<input type="hidden" name="phone"				value="<%=phone%>">
			<!--
			<input type="hidden" name="demo"				value="Y">
			-->
			<br>
			<center>
				<b><font color=red><%=langGenPayNowMsg%></font></b>
				<br><br>
				<b><font color=red size=2>--&gt;&nbsp;&nbsp;&nbsp;</font></b>
				<input type="image" src="../UserMods/butt_2CheckOut.gif" border="0" name="submit" alt="Submit Payment" align="middle">
				<b><font color=red size=2>&nbsp;&nbsp;&nbsp;&lt;--</font></b>
			</center>
			<br>
		</TD>
	</TR>
	</form>
<%
end sub
'**********************************************************************
'AuthorizeNet WebLink payments
'Notes : Always returns control to the same URL, regardless of the 
'        status of the transaction. The return URL is passed to the 
'        Authorize.Net routine. When control is returned, Authorize.Net
'        passes a status indicator which can be checked.
'**********************************************************************
sub payAuthorizeNet()

	'Check if we are in demo mode
	if demoMode = "Y" then
%>
	<form method="post" action="<%=urlSsl%>60_PayReturn.asp?CP_idOrder=<%=qIDOrder%>&CP_Status=success" id=AuthorizeNetForm name=AuthorizeNetForm>
<%
	else
%>
	<form method="post" action="https://secure.authorize.net/gateway/transact.dll" id=AuthorizeNetForm name=AuthorizeNetForm>
<%
	end if
%>
	<TR> 
		<TD class="CPgenHeadings">
			<%call InsertFP(authNetLogin,authNetTxKey,moneyD(total),qIdOrder,authNetCurrCode)%>
			<input type="hidden" name="x_version"				value="3.1">
			<input type="hidden" name="x_type"					value="AUTH_CAPTURE">
			<input type="hidden" name="x_Show_Form"				value="PAYMENT_FORM">
			<input type="hidden" name="x_method"				value="CC">
			<input type="hidden" name="x_Email_Customer"		value="TRUE">
			<input type="hidden" name="x_Email_Merchant"		value="TRUE">
			<input type="hidden" name="x_Login"					value="<%=authNetLogin%>">
			<input type="hidden" name="x_Amount"				value="<%=moneyD(total)%>">
			<input type="hidden" name="x_Invoice_Num"			value="<%=qIdOrder%>">
			<input type="hidden" name="x_Description"			value="<%=pCompany & " Order " & pOrderPrefix & "-" & qIdOrder%>">
			<input type="hidden" name="x_currency_code"			value="<%=authNetCurrCode%>">
			<input type="hidden" name="x_cust_id"				value="<%=idCust%>">
			<input type="hidden" name="x_first_name"			value="<%=name%>">
			<input type="hidden" name="x_last_name"				value="<%=Lastname%>">
			<input type="hidden" name="x_address"				value="<%=address%>">
			<input type="hidden" name="x_city"					value="<%=city%>">
			<input type="hidden" name="x_zip"					value="<%=zip%>">
			<input type="hidden" name="x_state"					value="<%=locState%>">
			<input type="hidden" name="x_country"				value="<%=locCountry%>">
			<input type="hidden" name="x_company"				value="<%=customerCompany%>">
			<input type="hidden" name="x_phone"					value="<%=phone%>">
			<input type="hidden" name="x_Email"					value="<%=email%>">
			<input type="hidden" name="x_Receipt_Link_URL"		value="<%=urlSsl%>60_PayReturn.asp">
			<input type="hidden" name="x_Receipt_Link_Method"	value="Post">
			<input type="hidden" name="x_Receipt_Link_Text"		value="Continue ...">
			<!--
			<input type="hidden" name="x_Test_Request"			value="TRUE">
			<input type="hidden" name="x_password"				value="testdriver">
			-->
			<br>
			<center>
				<b><font color=red><%=langGenPayNowMsg%></font></b>
				<br><br>
				<b><font color=red size=2>--&gt;&nbsp;&nbsp;&nbsp;</font></b>
				<input type="image" src="../UserMods/butt_AuthNet.gif" border="0" name="submit" alt="Submit Payment" align="middle">
				<b><font color=red size=2>&nbsp;&nbsp;&nbsp;&lt;--</font></b>
			</center>
			<br>
		</TD>
	</TR>
	</form>
<%
end sub
'**********************************************************************
'Custom payments
'Notes : Custom payments should only be used if the appropriate code 
'      : has been entered into the custom payment user include files.
'**********************************************************************
sub payCustomRoutine()
%>
	<TR> 
		<TD class="CPgenHeadings">
<!--#include file="../UserMods/_INCpayOut_.asp"-->
		</TD>
	</TR>
<%
end sub
'*************************************************************************
'Get Country Code from Country Description
'*************************************************************************
function getCountryCode(locName)

	dim mySQL, rsTemp
	
	getCountryCode = trim(locName)
	
	'Get Country Code
	mySQL = "SELECT locCountry " _
	      & "FROM   locations " _
	      & "WHERE  locName = '" & validSQL(trim(locName),"A") & "' " _
	      & "AND   (locState = '' OR locState IS NULL)"
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.eof then
		getCountryCode = rsTemp("locCountry")
	end if
	call closeRS(rsTemp)
	
end function
'*************************************************************************
'Get State Code from State Description and Country Code
'*************************************************************************
function getStateCode(locName,countryCode)

	dim mySQL, rsTemp
	
	getStateCode = trim(locName)
	
	'Get State Code
	mySQL = "SELECT locState " _
	      & "FROM   locations " _
	      & "WHERE  locName = '"    & validSQL(trim(locName),"A")     & "' " _
	      & "AND    locCountry = '" & validSQL(trim(countryCode),"A") & "' " _
	      & "AND    NOT(locState = '' OR locState IS NULL)"
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.eof then
		getStateCode = rsTemp("locState")
	end if
	call closeRS(rsTemp)
		
end function
%>
