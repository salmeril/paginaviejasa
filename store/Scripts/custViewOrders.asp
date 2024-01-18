<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Display order details from customer's account
'          : Page can also be called as last step of checkout process
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
'cartHead
dim randomKey
dim orderStatus
dim orderDate
dim subTotal
dim taxTotal
dim shipmentTotal
dim handlingFeeTotal
dim Total
dim shipmentMethod
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
dim shippingName
dim shippingLastName
dim shippingPhone
dim shippingAddress
dim ShippingCity
dim shippingZip
dim shippingLocState
dim shippingLocCountry
dim paymentType
dim cardType
dim cardNumber
dim cardExpMonth
dim cardExpYear
dim cardName
dim cardVerify
dim generalComments
dim deliverydate
dim storeComments
dim adjustReason
dim adjustAmount
dim discCode
dim discPerc
dim discTotal

'cartRows
dim IDCartRow
dim IDProduct
dim SKU
dim Quantity
dim unitPrice
dim Description
dim discAmt

'cartRowsOptions
dim idOption
dim optionPrice
dim optionDescrip

'Products
dim fileName

'DiscProd
dim idDiscProd
dim discFromQty
dim discToQty

'Work Fields
dim f
dim qIdOrder
dim optionGroupsTotal
dim optionsDisplay
dim refererURL
dim action

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

'Check that the Customer logged in
if isNull(idCust) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrNotLoggedIn)
end if

'Validate Order Number from QueryString
qIdOrder = Request.QueryString("idOrder")
if len(qIdOrder) = 0 or not IsNumeric(qIdOrder) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
end if

'Retrieve some information we need from cartHead
mySQL="SELECT randomKey,orderStatus,orderDate,subTotal,taxTotal," _
	& "       shipmentTotal,Total,shipmentMethod,Name,LastName," _
	& "       CustomerCompany,Phone,Email,Address,City,locState," _
	& "       locCountry,Zip,shippingName,shippingLastName,shippingPhone," _
	& "       shippingAddress,ShippingCity,shippingLocState," _
	& "       shippingLocCountry,shippingZip,paymentType,cardType," _
	& "       cardNumber,cardExpMonth,cardExpYear,cardVerify," _
	& "       cardName,generalComments,deliverydate,adjustReason,adjustAmount," _
	& "       discCode,discPerc,discTotal,handlingFeeTotal," _
	& "       storeComments " _
	& "FROM   cartHead " _
	& "WHERE  idOrder = " & validSQL(qIdOrder,"I") & " " _
	& "AND    idCust = "  & validSQL(idCust,"I")
set rsTemp = openRSexecute(mySQL)
if not rstemp.eof then

	'Assign to local variables
	storeComments		= rstemp("storeComments")
	randomKey           = rstemp("randomKey")
	orderStatus			= rstemp("orderStatus")
	orderDate			= rstemp("orderDate")
	subTotal			= rstemp("subTotal")
	taxTotal			= rstemp("taxTotal")
	shipmentTotal		= rstemp("shipmentTotal")
	Total				= rstemp("Total")
	shipmentMethod		= rstemp("shipmentMethod")
	Name				= rstemp("name")
	LastName			= rstemp("LastName")
	CustomerCompany		= rstemp("CustomerCompany")
	Phone				= rstemp("Phone")
	Email				= rstemp("Email")
	Address				= rstemp("Address")
	City				= rstemp("City")
	Zip					= rstemp("Zip")
	locState			= rstemp("locState")
	locCountry			= rstemp("locCountry")
	shippingName		= rstemp("shippingName")
	shippingLastName	= rstemp("shippingLastName")
	shippingPhone		= rstemp("shippingPhone")
	shippingAddress		= rstemp("shippingAddress")
	ShippingCity		= rstemp("ShippingCity")
	shippingZip			= rstemp("shippingZip")
	shippingLocState	= rstemp("shippingLocState")
	shippingLocCountry	= rstemp("shippingLocCountry")
	paymentType			= rstemp("paymentType")
	cardType			= rstemp("cardType")
	cardNumber			= rstemp("cardNumber")
	cardExpMonth		= rstemp("cardExpMonth")
	cardExpYear			= rstemp("cardExpYear")
	cardName			= rstemp("cardName")
	cardVerify			= rstemp("cardVerify")
	generalComments		= rstemp("generalComments")
	deliverydate		= rstemp("deliverydate")
	adjustReason		= rstemp("adjustReason")
	adjustAmount		= rstemp("adjustAmount")
	discCode			= rstemp("discCode")
	discPerc			= rstemp("discPerc")
	discTotal			= rstemp("discTotal")
	handlingFeeTotal	= rstemp("handlingFeeTotal")
	
	'Decrypt Card Number (if required)
	cardNumber = EnDeCrypt(Hex2Ascii(cardNumber),rc4Key)
	
	'Cater for orders entered before order discounts were added
	if isNull(discPerc) then
		discPerc = 0.00
	end if
	if isNull(discTotal) then
		discTotal = 0.00
	end if
	
else
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
end if
call closeRS(rsTemp)

'What page did we come from?
refererURL = lCase(Request.ServerVariables("HTTP_REFERER"))

'Check if we are showing the 'print' version of the page
action = trim(Request.QueryString("action"))
if lCase(action) <> "print" then
	action = ""
end if

'If we are NOT printing this page, display the full header
if action <> "print" then
%>
	<!--#include file="../UserMods/_INCtop_.asp"-->
<%
'If we are printing this page, show bare-bones header
else
%>
	<html>
	<head>
		<title>Invoice</title>
		<style type="text/css">
		<!--
		BODY, B, TD, P {COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt;}
		.CPlines {BACKGROUND-COLOR: #CCCC99}
		.CPgenHeadings {BACKGROUND-COLOR: #EEEEEE}
		-->
		</style>
	</head>
	<body>
	<table border=1 cellpadding=5 cellspacing=0 bgColor="#FFFFCC">
		<tr><td><b><%=langGenPrintMessage%></b></td></tr>
	</table>
<%
end if
%>

<!-- Outer Table Cell -->
<table border="0" cellpadding="0" cellspacing="0" width=450><tr><td>

<!-- Heading, Links and Pending Message -->
<%
if action <> "print" then
	call pageHeading()
	call defaultLinks()
	call pendingMessage()
end if
%>

<!-- Order Summary -->
<table border="0" cellpadding="1" cellspacing="2">
	<tr>
		<td valign=top nowrap><b><%=langGenOrderNumber%></b>&nbsp;</td>
		<td valign=top><%=pOrderPrefix & "-" & qIdOrder%></td>
	</tr>
	<tr>
		<td valign=top nowrap><b><%=langGenOrderDate%></b>&nbsp;</td>
		<td valign=top><%=formatTheDate(orderDate)%></td>
	</tr>
	<tr>
		<td valign=top nowrap><b><%=langGenOrderStatus%></b>&nbsp;</td>
		<td valign=top><font color=red><%=orderStatusDesc(orderStatus)%></font></td>
	</tr>
	<tr>
		<td valign=top nowrap><b><%=langGenStoreComments%></b>&nbsp;</td>
		<td valign=top><%=replace(emptyString(storeComments,"None"),Chr(10),"<br>")%></td>
	</tr>
</table>

<%call drawHLine()%>
   
<!-- Billing and Shipping -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr>
		<td class="CPgenHeadings">&nbsp;</td>
		<td width="50%" class="CPgenHeadings"><b><%=langGenBillAddr%></b></td>
		<td width="50%" class="CPgenHeadings"><b><%=langGenShipAddr%></b></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenName%></b>&nbsp;</td>
		<td><%=Name & " " & LastName%></td>
		<td><%=emptyString(shippingName,Name) & " " & emptyString(shippingLastName,LastName)%></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenAddress%></b>&nbsp;</td>
		<td><%=address%></td>
		<td><%=emptyString(shippingAddress,address)%></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenCity%></b>&nbsp;</td>
		<td><%=city%></td>
		<td><%=emptyString(shippingCity,city)%></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenZip%></b>&nbsp;</td>
		<td><%=zip%></td>
		<td><%=emptyString(shippingZip,zip)%></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenLocation%></b>&nbsp;</td>
		<td>
<%
			if len(locState) > 0 then
				Response.Write locState & ", "
			end if
			Response.Write locCountry
%>
		</td>
		<td>
<%
			if len(shippingLocState) = 0 and len(shippingLocCountry) = 0 then
				if len(locState) > 0 then
					Response.Write locState & ", "
				end if
				Response.Write locCountry
			else
				if len(shippingLocState) > 0 then
					Response.Write shippingLocState & ", "
				end if
				Response.Write shippingLocCountry
			end if
%>
		</td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenPhone%></b>&nbsp;</td>
		<td><%=Phone%></td>
		<td><%=emptyString(shippingPhone,Phone)%></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenCompany%></b>&nbsp;</td>
		<td colspan=2><%=CustomerCompany%></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenShipping%></b>&nbsp;</td>
		<td colspan=2><i><%=shipmentMethod%></i></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenPayment%></b>&nbsp;</td>
		<td colspan=2><i><%=paymentMsg(paymentType, total, cardNumber)%></i></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGenComments%></b>&nbsp;</td>
		<td colspan=2><i><%=emptyString(generalComments,"None")%></i></td>
	</tr>
	<tr>
		<td nowrap class="CPgenHeadings"><b><%=langGendeliverydate%></b>&nbsp;</td>
		<td colspan=2><i><%=emptyString(deliverydate,"None")%></i></td>
	</tr>
</table>

<!-- Order Details & Totals -->
<a name="orderItems">
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr> 
		<td width="10%" class="CPgenHeadings" nowrap><b><%=langGenQty%></b></td>
		<td width="80%" class="CPgenHeadings" nowrap><b><%=langGenItemDesc%></b></td>
		<td width="10%" class="CPgenHeadings" nowrap><b><%=langGenSubTotal%></b></td>
	</tr>
<%
	'Get all rows for this cart
	mySQL = "SELECT idCartRow,idProduct,quantity," _
	      & "       unitPrice,description,sku,discAmt " _
	      & "FROM   cartRows " _
		  & "WHERE  cartRows.idOrder = " & validSQL(qIdOrder,"I") & " " _
		  & "ORDER BY description"
	set rsTemp = openRSexecute(mySQL)
	do while not rstemp.eof

		'Assign record values to local values
		IDCartRow	= rstemp("idCartRow")
		IDProduct	= rstemp("idProduct")
		Quantity	= rstemp("quantity")
		unitPrice	= rstemp("unitPrice")
		Description	= rstemp("description")
		SKU			= rstemp("sku")
		discAmt		= rstemp("discAmt")
		
		'Cater for orders entered before discounts were added
		if isNull(discAmt) then
			discAmt = 0.00
		end if
%> 
		<tr> 
			<td nowrap valign=top><%=Quantity%></td>
			<td valign=top>
<%
				if SKU = "" then 
					'Use idProduct
					response.write IDProduct
				else
					'Use sku
					response.write SKU
				end if
				
				'Write cartRow line (main item)
				response.write " " & Description & " - <i>" & pCurrencySign & moneyS(unitPrice) & "</i> "
					
				'Display Download link if required, unless this is a 
				'printable view of the page in which case it's ignored.
				if action <> "print" then
					fileName = downloadFile(qIdOrder,IDCartRow)
					if fileName <> "" then
						Response.Write " (<a href=""" & urlNonSSL & "sysDownload.asp?randomKey=" & randomKey & "&idOrder=" & qIdOrder & "&idCartRow=" & IDCartRow & """>" & langGenDownload & "</a>)"
					end if
				end if
				Response.Write "<br>"
				
				'Write Discount (if any)
				if discAmt > 0 then
					Response.Write "* <i>" & langGenDiscount & " - " & pCurrencySign & moneyS(discAmt) & "</i><br>"
				end if
				
				'Get all options for this row
				optionGroupsTotal = 0
				mySQL = "SELECT optionPrice, optionDescrip " _
					  & "FROM   cartRowsOptions " _
					  & "WHERE  idCartRow = " & validSQL(idCartRow,"I")
				set rsTemp2 = openRSexecute(mySQL)
				do while not rstemp2.eof
					
					'Assign record values to local values
					optionDescrip = rstemp2("optionDescrip")
					optionPrice	  = rstemp2("optionPrice")
						
					'Write cartRowOptions line(s) (options)
					Response.Write "* <i>" & optionDescrip
					if optionPrice <> 0 then
						Response.Write " - " & pCurrencySign & moneyS(optionPrice)
					end if
					Response.Write "</i><br>"

					'Calculate options Sub Total
					optionGroupsTotal = optionGroupsTotal + optionPrice        
					rstemp2.movenext
				loop
				call closeRS(rsTemp2)
%>
			</td>
			<td nowrap valign=top>
				<%=pCurrencySign &  moneyS(Cdbl(Quantity * (optionGroupsTotal + unitPrice - discAmt)))%> 
			</td>
		</tr>
<%
		rstemp.movenext
	loop
	
	call closeRS(rsTemp)
%>
	<tr> 
		<td colspan="2" align=right class="CPgenHeadings">
			<b><%=langGenSubTotal%>:&nbsp;&nbsp;</b>
		</td>
		<td align=left class="CPgenHeadings" nowrap> 
			<b><%=pCurrencySign & moneyS(subTotal)%></b>
		</td>
	</tr>
	<tr> 
		<td colspan="2" align=right>
<%
			if discTotal > 0 then
				Response.Write "<i>" & discCode & " (" & formatNumber(discPerc,2) & "%)" & "</i> - "
			end if
%>
			<b><%=langGenDiscCode%>:&nbsp;&nbsp;</b>
		</td>
		<td align=left nowrap> 
<%
			Response.Write pCurrencySign & moneyS(discTotal)
			if discTotal > 0 then
				Response.Write "&nbsp;&nbsp;(-)"
			end if
%>
		</td>
	</tr>
	<tr> 
		<td colspan="2" align=right>
			<b><%=langGenSubTotal%>:&nbsp;&nbsp;</b>
		</td>
		<td align=left nowrap> 
			<b><%=pCurrencySign & moneyS(subTotal - discTotal)%></b>
		</td>
	</tr>
	<tr> 
		<td colspan="2" align=right>
			<b><%=langGenShipping%>:&nbsp;&nbsp;</b>
		</td>
		<td align=left nowrap> 
			<%=pCurrencySign & moneyS(shipmentTotal)%>
		</td>
	</tr>
<%	if discTotal > 0 then
%>	<tr> 
		<td colspan="2" align=right>
			<b><%=langGenHandlingFee%>:&nbsp;&nbsp;</b>
		</td>
		<td align=left nowrap> 
			<%=pCurrencySign & moneyS(handlingFeeTotal)%>
		</td>
	</tr>
<%	end if
%>	<tr> 
		<td colspan="2" align=right>
			<b><%=langGenTax%>:&nbsp;&nbsp;</b>
		</td>     
		<td align=left nowrap> 
			<%=pCurrencySign & moneyS(taxTotal)%>&nbsp;
		</td>
	</tr>
	<tr> 
		<td colspan="2" align=right>
			*<b><%=langGenAdjustment%>:&nbsp;&nbsp;</b>
		</td>     
		<td align=left nowrap>
<%
			if isNumeric(adjustAmount) then
				Response.Write pCurrencySign & moneyS(adjustAmount)
			else
				Response.Write pCurrencySign & moneyS("0.00")
			end if
%>
		</td>
	</tr>
	<tr>
		<td colspan="2" align=right class="CPgenHeadings">
			<b><%=langGenTotal%>:&nbsp;&nbsp;</b>
		</td>
		<td align=left class="CPgenHeadings" nowrap> 
			<b><%=pCurrencySign & moneyS(Total)%></b>
		</td>
	</tr>
	<tr>
		<td colspan="3">
			*<b><%=langGenAdjustment%> :</b> 
<%
			if len(adjustReason) > 0 then
				Response.Write adjustReason
			else
				Response.Write langGenNotApplicable
			end if
%>
		</td>
	</tr>
	<tr><td colspan="3">&nbsp;</td></tr>
	
</table>

<!-- End Outer Table Cell -->
</TD></TR></TABLE>

<br><br>

<%
'If we are NOT printing this page, display full footer
if action <> "print" then
%>
	<!--#include file="../UserMods/_INCbottom_.asp"-->
<%
'If we are printing this page, show bare-bones footer
else
%>
	</body></html>
<%
end if

'Close Database connection
call closeDB()

'**********************************************************************
'Draw horizontal line
'**********************************************************************
sub drawHLine()
%>
	<img border="0" height="4" width="1" src="../UserMods/misc_cleardot.gif"><br>
	<table border=0 cellpadding=0 cellspacing=0 width="100%">
		<tr><td width="100%" class="CPlines">
			<img border="0" height="1" width="1" src="../UserMods/misc_cleardot.gif"><br>
		</td></tr>
	</table>
	<img border="0" height="4" width="1" src="../UserMods/misc_cleardot.gif"><br>
<%
end sub
'**********************************************************************
'Page Heading
'**********************************************************************
sub pageHeading()
%>
	<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr><td nowrap valign=middle class="CPpageHead">
		<table border=0 cellpadding=0 cellspacing=0 width="100%">
		<tr>
<%
		'Thank You heading
		if instr(refererURL,lCase("40_SubmitOrder.asp")) <> 0 then
%>
			<td nowrap valign=middle>
				<b><%=langGenPaySuccessHdr%></b>
			</td>
			<td nowrap align=right valign=middle>
				<b><font color=#800000>[ <%=langGenStep%> 4 / 4 ]</font></b>
			</td>
<%
		'View Order heading
		else
%>
			<td nowrap valign=middle>
				<b><%=langGenOrderView%></b>
			</td>
<%
		end if
%>
		</tr>
		</table>
	</td></tr>
	</table>
	<img src="../UserMods/misc_cleardot.gif" height=2 width=1><br>
<%
end sub
'**********************************************************************
'Display default links
'**********************************************************************
sub defaultLinks()
%>
	<table border="0" cellpadding="2" cellspacing="0" width="100%">
		<tr>
			<td nowrap valign=middle>
				&raquo;&nbsp;<a href="custListOrders.asp"><%=langGenYourAccount%></a>&nbsp;&nbsp;
<%
				'Check if we must show Download link
				if orderHasDownloads(qIdOrder) then 
%>
					&raquo;&nbsp;<a href="#orderItems"><%=langGenDownload%></a></b>&nbsp;&nbsp;
<%
				end if 
%>
				&raquo;&nbsp;<a href="custViewOrders.asp?idOrder=<%=qIdOrder%>&action=print" target="_blank"><%=langGenPrintVersion%></a>&nbsp;&nbsp;
				&raquo;&nbsp;<a href="<%=urlNonSSL%>termsAndCond.asp" onClick='window.open("<%=urlNonSSL%>termsAndCond.asp","generalConditions","width=300,height=300,resizable=1,scrollbars=1");return false;' target="_blank"><%=langGenPayPolicy%></a><br>
			</td>
		</tr>
	</table>
<%
	call drawHLine()
end sub
'**********************************************************************
'Show payment pending message if order status is pending
'**********************************************************************
sub pendingMessage()
	If orderStatus = "0" then
%>
		<table border="0" cellpadding="2" cellspacing="0" width="100%">
			<tr>
				<td valign=middle class="CPgenHeadings">
<%
				'Write "Payment Pending" message
				Response.Write langGenOrdPendingMsg
				
				'Give the customer the opportunity to re-submit 
				'payment for certain payment types.
				if lCase(paymentType)="paypal" _
				or lCase(paymentType)="2checkout" _
				or lCase(paymentType)="authorizenet" _
				or lCase(paymentType)="custom" then
					'Order Number is passed via the session object for the 
					'benefit of gateways that require the payment page to be 
					'a fixed URL.
					session(storeID & "idOrderPaySubmit") = qIdOrder
%>
					&nbsp;&nbsp;<a href="50_PaySubmit.asp"><%=langGenReSubPay%></a>
<%
				end if
%>
				</td>
			</tr>
		</table>
<%
		call drawHLine()
	end if
end sub
'******************************************************************
'Check if an Order has any Downloadable Items
'******************************************************************
function orderHasDownloads(idOrder)

	if isEmpty(idOrder) or not IsNumeric(idOrder) then
		orderHasDownloads = false
		exit function
	end if

	dim mySQL, rsTemp
	mySQL="SELECT cartRows.idProduct " _
		& "FROM   cartRows, products " _
		& "WHERE  idOrder = " & validSQL(idOrder,"I") & " " _
		& "AND    products.idProduct = cartRows.idProduct " _
		& "AND    NOT (products.fileName IS NULL " _
		& "OR     products.fileName = '') "
	set rsTemp = openRSexecute(mySQL)
	if rsTemp.eof then
		orderHasDownloads = false
	else
		orderHasDownloads = true
	end if
	call closeRS(rsTemp)
end function
%>
