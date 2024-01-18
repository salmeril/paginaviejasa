<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : 2CheckOut.Com auto order update.
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
'1. The order is automatically updated if it passes all the tests.
'2. This script is called upon return from 2CheckOut.Com
'3. This script is a lot stricter than the generic "Thank You" page 
'   due to the auto update. All fields MUST be present and valid in 
'   the context of the order. If not, no update is performed and the 
'   customer will see an "Error with Payment" type of message.
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
<!--#include file="_INCmd5_.asp"-->
<%
'Work Fields
dim qIdOrder
dim qIdOrder2CO
dim qTotal
dim qKey
dim statusInd
dim formattedDateTime
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

'Get a Date and Time formatted to the user's specifications
formattedDateTime = formatTheDate(now()) & " " & time()

'Get Form variables
qIdOrder	= trim(Request.Form("cart_order_id"))
qIdOrder2CO	= trim(Request.Form("order_number"))
qTotal		= trim(Request.Form("total"))
qKey		= trim(Request.Form("key"))
statusInd	= trim(Request.Form("credit_card_processed"))

'Validate Form variables were passed
if len(qIdOrder)=0 or len(qIdOrder2CO)=0 or len(qTotal)=0 or len(qKey)=0 then
	statusInd = "error"
else

	'Check status passed back by 2CheckOut.Com
	if UCase(statusInd) = "Y" then
	
		'Check MD5 Hash Key
		if UCase(qKey) = UCase(md5(TwoCheckoutMD5 & TwoCheckOutSID & qIdOrder2CO & qTotal)) then
		
			'Check the current Order Status for "Pending"
			mySQL = "SELECT orderStatus " _
				  & "FROM   cartHead " _
				  & "WHERE  idOrder = " & validSQL(qIdOrder,"I")
			set rsTemp = openRSexecute(mySQL)
			if rsTemp.eof then
				statusInd = "error"
			else
				statusInd = "success"
				'Note : We show a "success" page even if the order
				'status is no longer "Pending". This is to avoid 
				'confusion where the customer clicks the button on 
				'the 2CheckOut.Com page which takes them to this 
				'page (which will update the order the first time), 
				'and then clicks the back button and click the 
				'2CheckOut button again.
				if rsTemp("orderStatus") = "0" then 
					call updOrderStatus(qIdOrder,"1","Y","Y","DATE : " & formattedDateTime & vbCrLf & "2CheckOut : Status = " & statusInd)
				end if
			end if
			call closeRS(rsTemp)

		else
			statusInd = "error"
		end if
	else
		statusInd = "error"
	end if
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
