<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Generic "Thank You" and "Error with Payment" page for 
'          : payments made via 3rd Party payment processors.
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
'1. No database updates are performed in this script. 
'2. This script is called upon return from a payment processor.
'3. This script is very "forgiving" in that it only shows the "Error 
'   with Payment" message if explicitly told to do so by the payment 
'   processor. If some of the return values are missing but no error 
'   was reported by the payment processor, the customer will still 
'   see a "Thank You" message, but some of the links will not be 
'   displayed.
'*************************************************************************
Option explicit
Response.Buffer = true
%>
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="_INCconfig_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="_INCappFunctions_.asp"-->
<%
'Work Fields
dim qIdOrder
dim statusInd
dim payMessage
dim payMessageVar

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
idCust  = sessionCust()

'Try to obtain the Order Number for the order in question. Different 
'gateways pass the Order Number back in different ways, thus we are 
'able to determine where we came from by checking what value was 
'returned. We also determine if this was a successful transaction or 
'not.

'--> PayPal
if len(qIdOrder) = 0 then

	'The 'item_number' value is 'supposed' to be passed back by 
	'PayPal, but we have found this to be problematic. This is why 
	'we constructed the return URL's to include the idOrder value, 
	'just in case 'item_number' doesn't work.
	qIdOrder = trim(Request.Form("item_number"))
	
	'If the above line didn't get the Order Number, try this one.
	if len(qIdOrder) = 0 then
		qIdOrder = trim(Request.QueryString("CP_idOrder"))
	end if
	
	'Get Status
	if len(qIdOrder) > 0 then
		statusInd = lCase(trim(Request.QueryString("CP_Status")))
	end if
	
end if

'--> 2Checkout
if len(qIdOrder) = 0 then

	'Get Order Number
	qIdOrder = trim(Request.Form("cart_order_id"))
	
	'Get Status
	if len(qIdOrder) > 0 then
		statusInd = trim(Request.Form("credit_card_processed"))
		if UCase(statusInd) = "Y" then
			statusInd = "success"
		else
			statusInd = "error"
		end if
	end if
	
end if

'--> Authorize.Net
if len(qIdOrder) = 0 then

	'Get Order Number
	qIdOrder = trim(Request.Form("x_invoice_num"))
	
	'Get Status
	if len(qIdOrder) > 0 then
		statusInd = trim(Request.Form("x_response_code"))
		if UCase(statusInd) = "1" then
			statusInd = "success"
		else
			statusInd = "error"
		end if
	end if
	
end if

'--> Custom Payment
%>
<!--#include file="../UserMods/_INCpayIn_.asp"-->
<%

'--> Set to zero length string if still incorrect
if len(qIdOrder) = 0 or not(IsNumeric(qIdOrder)) then
	qIdOrder  = ""
	statusInd = ""
end if

%> 

<!--#include file="../UserMods/_INCtop_.asp"-->

<!-- Outer Table Cell -->
<table border="0" cellpadding="0" cellspacing="0" width=450><tr><td>

<!-- Heading -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr><td nowrap valign=middle class="CPpageHead">
<%
		if statusInd = "error" then
%>
			<b><font color=red><%=langGenPayErrorHdr%></font></b>
<%
		else
%>
			<b><%=langGenPaySuccessHdr%></b>
<%
		end if
%>
	</td></tr>
</table>
	
<br>

<!-- Links -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr><td valign=middle class="CPgenHeadings">
		&raquo;&nbsp;<b><a href="<%=urlSSL%>custListOrders.asp"><%=langGenYourAccount%></a></b>&nbsp;&nbsp;&nbsp;&nbsp;
<%
		'If Order Number is available, link to it
		if len(qIdOrder) > 0 and IsNumeric(qIdOrder) then
%>
			&raquo;&nbsp;<b><a href="<%=urlSSL%>custViewOrders.asp?idOrder=<%=qIdOrder%>"><%=langGenOrder & " " & pOrderPrefix & "-" & qIdOrder%></a></b>
<%	
		end if
%>
	</td></tr>
</table>

<br>

<!-- Message -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr><td>
<%
		'Get appropriate payment message from database
		if statusInd = "error" then
			payMessageVar = "payErrorMsg"
		else
			payMessageVar = "paySuccessMsg"
		end if
		mySQL = "SELECT configValLong " _
			&   "FROM   storeAdmin " _
			&   "WHERE  configVar='" & payMessageVar & "' " _
			&   "AND    adminType='T'"
		set rsTemp = openRSexecute(mySQL)
		if not rstemp.eof then
			payMessage = trim(rsTemp("configValLong"))
		end if
		call closeRS(rsTemp)
	
		'Check for tags and replace
		payMessage = replace(payMessage,"#STORE#",pCompany)
		payMessage = replace(payMessage,"#SALES#","<a href=""mailto:" & pEmailSales & """>" & pEmailSales & "</a>")
			
		'Display Message
		Response.Write payMessage
%>
	</td></tr>
</table>

<!-- End Outer Table Cell -->
</td></tr></table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%
call closeDB()
%>
