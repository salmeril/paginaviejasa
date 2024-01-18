<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Calculate tax and other totals
'          : Show complete order for final user approval/disapproval
'		   : Cancel Order
'		   : Submits Order
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
<!--#include file="_INCappEmail_.asp"-->
<!--#include file="_INCupdStatus_.asp"-->
<%
'cartHead
dim orderDate
dim orderStatus
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
dim taxExempt
dim discCode
dim discPerc
dim discTotal

'customer
dim locState2
dim shippingLocState2

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

'DiscProd
dim idDiscProd
dim discFromQty
dim discToQty

'Work Fields
dim f
dim customerEmail
dim customerEmailItems
dim optionGroupsTotal
dim optionGroupsDesc

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

'Check if the session is still active
if isNull(idOrder) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrCartEmpty)
end if

'Check if cart has any items
if cartQty(idOrder) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrCartEmpty)
end if

'Check if minimum order amount has been met
if cartTotal(idOrder,0) < pMinCartAmount then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrMinPrice & pCurrencySign & moneyS(pMinCartAmount))
end if

'Check that the Customer is logged on
if isNull(idCust) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrNotLoggedIn)
end if

'Check if "Cancel Order" button was clicked
if request.form("CancelOrder") <> "" then

	'Cancel Order
	CancelOrder()
	
	'Close DB Connection
	call closeDB()
	
	'Go to index page
	response.redirect urlNonSSL & "default.asp"
	
end if

'Check if "Submit Order" button was clicked
if request.form("SubmitOrder") <> "" then

	'Submit the Order
	SubmitOrder()
	
	'Close Database
	call closeDb()
	
	'Check Order Status and Payment Type
	if orderStatus = "0" then	'Pending
		select case lCase(paymentType)
		case "paypal","2checkout","authorizenet","custom"
			'Order Number is passed via the session object for the 
			'benefit of gateways that require the payment page to be 
			'a fixed URL.
			session(storeID & "idOrderPaySubmit") = idOrder
			Response.Redirect "50_PaySubmit.asp"
		case else
			Response.Redirect "custViewOrders.asp?idOrder=" & idOrder 
		end select
	else
		Response.Redirect "custViewOrders.asp?idOrder=" & idOrder 
	end if
	
end if

'If we get this far, then neither "Cancel Order" nor "Submit Order" 
'was pressed. This means that we have not yet showed the Order for 
'viewing, so we do that here.

'Retrieve info from Customer and cartHead
getCustCartInfo()

'Calculate Tax and other Totals
taxAndTotals()
		
'Update cartHead with Totals and info from Customer
updateCartInfo()	

'Make sure this page can not be cached
Response.Expires = -10000
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

%>
<!--#include file="../UserMods/_INCtop_.asp"-->
<%

'Display the Order
showOrder()

%>
<!--#include file="../UserMods/_INCbottom_.asp"-->
<%

'Close the Database
call closeDB()		

'**********************************************************************
'Display the current Order
'**********************************************************************
sub showOrder()
%> 
	<!-- Outer Table Cell -->
	<table border="0" cellpadding="0" cellspacing="0" width=450><tr><td>
	
	<!-- Heading -->
	<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr><td nowrap valign=middle class="CPpageHead">
		<table border=0 cellpadding=0 cellspacing=0 width="100%">
		<tr>
			<td nowrap valign=middle>
				<b><%=langGenVerifyOrder%></b>
			</td>
			<td nowrap align=right valign=middle>
				<b><font color=#800000>[ <%=langGenStep%> 3 / 4 ]</font></b>
			</td>
		</tr>
		</table>
	</td></tr>
	</table>
	
	<img src="../UserMods/misc_cleardot.gif" height=4 width=1><br>
	
	<!-- Billing and Shipping -->
	<table border="0" cellspacing="0" cellpadding="2" width="100%">
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
				'If there is a state or alternate state, display it
				if len(trim(locState)) > 0 or len(trim(locState2)) > 0 then
					Response.Write getStateDesc(locCountry,locState,locState2) & ", "
				end if
				'Display country
				Response.Write getCountryDesc(locCountry)
%>
			</td>
			<td>
<%
				'If shipping state/country is blank, use billing state/country
				if len(shippingLocCountry) = 0 and len(shippingLocState) = 0 and len(shippingLocState2) = 0 then
					'If there is a state or alternate state, display it
					if len(trim(locState)) > 0 or len(trim(locState2)) > 0 then
						Response.Write getStateDesc(locCountry,locState,locState2) & ", "
					end if
					'Display country
					Response.Write getCountryDesc(locCountry)
				else
					'If there is a state or alternate state, display it
					if len(trim(shippingLocState)) > 0 or len(trim(shippingLocState2)) > 0 then
						Response.Write getStateDesc(shippingLocCountry,shippingLocState,shippingLocState2) & ", "
					end if
					'display country
					Response.Write getCountryDesc(shippingLocCountry)
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
	<TABLE BORDER="0" CELLPADDING="2" cellspacing="0" width="100%">
		<tr> 
			<td width="10%" class="CPgenHeadings" nowrap><b><%=langGenQty%></b></td>
			<td width="80%" class="CPgenHeadings" nowrap><b><%=langGenItemDesc%></b></td>
			<td width="10%" class="CPgenHeadings" nowrap><b><%=langGenSubTotal%></b></td>
		</tr>
<%
		'Get all rows for this cart
		mySQL = "SELECT idCartRow,idProduct,quantity," _
		      & "       unitPrice,description,sku," _
		      & "       discAmt " _
		      & "FROM   cartRows " _
			  & "WHERE  cartRows.idOrder = " & validSQL(idOrder,"I") & " " _
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
			if isNull(discAmt) then
				discAmt = 0.00
			end if
%> 
			<tr> 
				<td nowrap valign=top><%=Quantity%></td>
				
				<td align=left valign=top>
<%
					if SKU = "" then 
						'Use idProduct
						response.write IDProduct
					else
						'Use sku
						response.write SKU
					end if
					
					'Write cartRow line (main item)
					response.write " " & Description & " - <i>" & pCurrencySign & moneyS(unitPrice) & "</i><br>"
					
					'Write Discount (if any)
					if discAmt > 0 then
						Response.Write "* <i>" & langGenDiscount & " - " & pCurrencySign & moneyS(discAmt) & "</i><br>"
					end if

					'Get all options for this row
					optionGroupsTotal = 0
					mySQL = "SELECT optionPrice, optionDescrip " _
						  & "FROM   cartRowsOptions " _
						  & "WHERE  idCartRow = " & validSQL(IDCartRow,"I")
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
<%		if handlingFeeTotal > 0 then%>
		<tr> 
			<td colspan="2" align=right>
				<b><%=langGenHandlingFee%>:&nbsp;&nbsp;</b>
			</td>
			<td align=left nowrap> 
				<%=pCurrencySign & moneyS(handlingFeeTotal)%>
			</td>
		</tr>
<%		end if%>
		
		
		<tr> 
			<td colspan="2" align=right>
				<b><%=langGenTax%>:&nbsp;&nbsp;</b>
			</td>     
			<td align=left nowrap> 
				<%=pCurrencySign & moneyS(taxTotal)%>&nbsp;
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
	</table>

	<!-- Bottom Rows -->
	<TABLE BORDER="0" CELLPADDING="0" cellspacing="0" width="100%">
		<tr>
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr>
			<form method="post" name="submit_order" action="40_SubmitOrder.asp">
			<td align=center width="50%"> 
				<input type="submit" name="SubmitOrder" value="<%=langGenSaveOrder%>">
			</td>
			<td align=center width="50%"> 
				<input type="submit" name="CancelOrder" value="<%=langGenCancelOrder%>">
			</td>
			</form>
		</tr>
		<tr>
			<td colspan="2">
				<br>
				<i><%=langGenTOSmsg%> <a href="<%=urlNonSSL%>termsAndCond.asp" onClick='window.open("<%=urlNonSSL%>termsAndCond.asp","generalConditions","width=300,height=300,resizable=1,scrollbars=1");return false;' target="_blank"><%=langGenTOSlink%></a></i><br><br>
			</td>
	    </tr>
	</table>
	    
	<!-- End Outer Table Cell -->
	</TD></TR></TABLE>

	<br><br>
<%
end sub
'**********************************************************************
'Retrieve information we already have in Customer & cartHead tables
'**********************************************************************
sub getCustCartInfo()
	'Retrieve information we already have in Customer table
	mySQL="SELECT idCust,Name,LastName,CustomerCompany,Phone,Email," _
	    & "       Address,City,Zip,locState,locState2,locCountry," _
	    & "       shippingName,shippingLastName,shippingPhone," _
	    & "       shippingAddress,ShippingCity,shippingZip," _
	    & "       shippingLocState,shippingLocState2," _
	    & "       shippingLocCountry,taxExempt " _
	    & "FROM   customer " _
	    & "WHERE  idCust = " & validSQL(idCust,"I")
	set rsTemp = openRSexecute(mySQL)
	if not rstemp.eof then
		Name				= trim(rstemp("name"))
		LastName			= trim(rstemp("LastName"))
		CustomerCompany		= trim(rstemp("CustomerCompany"))
		Phone				= trim(rstemp("Phone"))
		Email				= trim(rstemp("Email"))
		Address				= trim(rstemp("Address"))
		City				= trim(rstemp("City"))
		Zip					= trim(rstemp("Zip"))
		locState			= trim(UCase(rstemp("locState")))
		locState2			= trim(rstemp("locState2"))
		locCountry			= trim(UCase(rstemp("locCountry")))
		shippingName		= trim(rstemp("shippingName"))
		shippingLastName	= trim(rstemp("shippingLastName"))
		shippingPhone		= trim(rstemp("shippingPhone"))
		shippingAddress		= trim(rstemp("shippingAddress"))
		ShippingCity		= trim(rstemp("ShippingCity"))
		shippingZip			= trim(rstemp("shippingZip"))
		shippingLocState	= trim(UCase(rstemp("shippingLocState")))
		shippingLocState2	= trim(rstemp("shippingLocState2"))
		shippingLocCountry	= trim(UCase(rstemp("shippingLocCountry")))
		taxExempt			= trim(UCase(rstemp("taxExempt")))
	end if
	call closeRS(rsTemp)

	'Retrieve some information we already have in Cart Header
	mySQL="SELECT shipmentMethod,shipmentTotal,paymentType,cardType," _
		& "       cardNumber,cardExpMonth,cardExpYear,cardName," _
		& "       cardVerify,generalComments,deliverydate,discCode,discPerc," _
		& "       handlingFeeTotal," _
	    & "      (SELECT paymentType " _
	    & "       FROM   customer " _
	    & "       WHERE  idCust = " & validSQL(idCust,"I") & ") " _
	    & "       AS     custPaymentType " _
	    & "FROM   cartHead " _
	    & "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	if not rstemp.eof then
		shipmentMethod   = trim(rstemp("shipmentMethod"))
		shipmentTotal    = rstemp("shipmentTotal")
		paymentType		 = trim(rstemp("paymentType"))
		cardType		 = trim(rstemp("cardType"))
		cardNumber		 = trim(EnDeCrypt(Hex2Ascii(rstemp("cardNumber")),rc4Key))
		cardExpMonth	 = trim(rstemp("cardExpMonth"))
		cardExpYear		 = trim(rstemp("cardExpYear"))
		cardName		 = trim(rstemp("cardName"))
		cardVerify		 = trim(rstemp("cardVerify"))
		generalComments	 = trim(rstemp("generalComments"))
		deliverydate	 = trim(rstemp("deliverydate"))
		discCode		 = trim(rstemp("discCode"))
		discPerc		 = trim(rstemp("discPerc"))
		handlingFeeTotal = rstemp("handlingFeeTotal")
		if isNull(discPerc) then
			discPerc = 0.00
		end if
	end if
	call closeRS(rsTemp)
end sub
'**********************************************************************
'Calculate Taxes and other Totals
'**********************************************************************
sub taxAndTotals()

	'Declare variables local to this function
	dim rsTemp, rsTemp2
	dim taxRate			'Tax rate to be applied to the order
	dim taxLocState		'Used to determine tax location
	dim taxLocCountry	'Used to determine tax location
	dim subTotalNoTax	'Sub total of non taxable items
	dim subTotalTaxable	'Sub total of taxable items
	dim subTotalTaxDisc 'Discount subtracted from taxable sub total
	dim unitPriceRow
	dim unitPriceOpt
	dim taxExemptRow
	dim taxExemptOpt
	
	'Initialize some variables
	subTotal        = 0.00
	discTotal       = 0.00
	taxTotal        = 0.00
	total           = 0.00
	subTotalNoTax   = 0.00
	subTotalTaxable = 0.00
	subTotalTaxDisc = 0.00
	
	'Determine what tax rate to use
	if taxExempt = "Y" then   'Customer is exempted from tax
		taxRate = 0.00
	else
		'Determine what location to use for tax purposes
		if taxBillOrShip = 0 then
			taxLocState	  = locState
			taxLocCountry = locCountry
		else
			if len(shippingLocCountry & shippingLocState) > 0 then
				taxLocState	  = shippingLocState
				taxLocCountry = shippingLocCountry
			else
				taxLocState	  = locState
				taxLocCountry = locCountry
			end if
		end if
		if len(locState) > 0 then	'By State/Province
			mySQL = "SELECT locTax " _
				  & "FROM   locations " _
				  & "WHERE  locCountry = '" & validSQL(taxLocCountry,"A") & "' " _
				  & "AND    locState = '"   & validSQL(taxLocState,"A")   & "' "
		else						'By Country
			mySQL = "SELECT locTax " _
				  & "FROM   locations " _
				  & "WHERE  locCountry = '" & validSQL(taxLocCountry,"A") & "' " _
				  & "AND   (locState = '' OR locState IS NULL)"
		end if
		set rsTemp = openRSexecute(mySQL)
		if rstemp.eof then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvTax)
		else
			taxRate = rstemp("locTax")
		end if
		call closeRS(rsTemp)
	end if
		
	'Read cartRows and calculate Sub Totals
	mySQL = "SELECT idCartRow,quantity,unitPrice," _
		  & "       taxExempt,discAmt " _
		  & "FROM   cartRows " _
		  & "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	do while not rsTemp.eof
	
		'Get values from cartRows
		idCartRow    = rsTemp("idCartRow")
		quantity     = rsTemp("quantity")
		unitPriceRow = rsTemp("unitPrice")
		taxExemptRow = UCase(trim(rsTemp("taxExempt")))
		discAmt      = rsTemp("discAmt")
		if isNull(discAmt) then
			discAmt = 0.00
		end if
		
		'Add item total to order SubTotal
		subTotal = subTotal + (quantity * (unitPriceRow - discAmt))
		
		'Add Row total to non-taxable SubTotal (if tax exempt)
		if taxExemptRow = "Y" then
			subTotalNoTax = subTotalNoTax + (quantity * (unitPriceRow - discAmt))
		end if
		
		'Read cartRowsOptions and calculate Sub Totals
		mySQL = "SELECT optionPrice,taxExempt " _
			  & "FROM   cartRowsOptions " _
			  & "WHERE  idCartRow = " & validSQL(idCartRow,"I")
		set rsTemp2 = openRSexecute(mySQL)
		do while not rsTemp2.eof
		
			'Get values from cartRowsOptions
			unitPriceOpt = rsTemp2("optionPrice")
			taxExemptOpt = UCase(trim(rsTemp2("taxExempt")))
			
			'Add options total to order SubTotal
			subTotal = subTotal + (quantity * unitPriceOpt)
			
			'Add options total to non-taxable SubTotal if exempt
			if taxExemptOpt = "Y" then
				subTotalNoTax = subTotalNoTax + (quantity * unitPriceOpt)
			end if

			rsTemp2.movenext
		loop
		call closeRS(rsTemp2)
		
		rsTemp.movenext
	loop
	call closeRS(rsTemp)
	
	'Calculate taxable sub total
	subTotalTaxable = subTotal - subTotalNoTax
	
	'Subtract order discount from taxable sub total
	subTotalTaxDisc = Round(((subTotalTaxable * discPerc) / 100),2)
	subTotalTaxable = subTotalTaxable - subTotalTaxDisc

	'Calculate tax total for taxable items
	if taxOnShipping = -1 then
		if handlingFeeTax = -1 then
			taxTotal = Round((((subTotalTaxable + shipmentTotal + handlingFeeTotal) * taxRate) / 100),2)
		else
			taxTotal = Round((((subTotalTaxable + shipmentTotal) * taxRate) / 100),2)
		end if
	else
		if handlingFeeTax = -1 then
			taxTotal = Round((((subTotalTaxable + handlingFeeTotal) * taxRate) / 100),2)
		else
			taxTotal = Round(((subTotalTaxable * taxRate) / 100),2)
		end if
	end if

	'Calculate total order discount
	discTotal = Round(((subTotal * discPerc) / 100),2)

	'Calculate final total
	Total = subTotal - discTotal + shipmentTotal + handlingFeeTotal + taxTotal
	
end sub
'**********************************************************************
'Update cartHead with remaining information
'**********************************************************************
sub updateCartInfo()
	'If we have come this far, we might as well update all the other fields
	'on CartHead in anticipation of the fact that the customer will go ahead 
	'and send the order to us. This makes "saving" the order a simple matter 
	'of updating the order status on the next page. 
	mySQL = "UPDATE cartHead SET " _
		&   "idCust				= "  & validSQL(idCust,"I") & ", " _
		&   "subTotal			= "  & validSQL(subTotal,"D") & ", " _
		&   "discTotal			= "  & validSQL(discTotal,"D") & ", " _
		&   "taxTotal			= "  & validSQL(taxTotal,"D") & ", " _
		&   "total				= "  & validSQL(total,"D") & ", " _
		&   "name				= '" & validSQL(name,"A") & "', " _
		&   "LastName			= '" & validSQL(LastName,"A") & "', " _
		&   "CustomerCompany	= '" & validSQL(CustomerCompany,"A") & "', " _
		&   "Phone				= '" & validSQL(Phone,"A") & "', " _
		&   "Email				= '" & validSQL(Email,"A") & "', " _
		&   "Address			= '" & validSQL(Address,"A") & "', " _
		&   "City				= '" & validSQL(City,"A") & "', " _
		&   "Zip				= '" & validSQL(Zip,"A") & "', " _
		&   "locState			= '" & validSQL(getStateDesc(locCountry,locState,locState2),"A") & "', " _
		&   "locCountry			= '" & validSQL(getCountryDesc(locCountry),"A") & "', " _
		&   "shippingName		= '" & validSQL(shippingName,"A") & "', " _
		&   "shippingLastName	= '" & validSQL(shippingLastName,"A") & "', " _
		&   "shippingPhone		= '" & validSQL(shippingPhone,"A") & "', " _
		&   "ShippingAddress	= '" & validSQL(ShippingAddress,"A") & "', " _
		&   "ShippingCity		= '" & validSQL(ShippingCity,"A") & "', " _
		&   "shippingZip		= '" & validSQL(shippingZip,"A") & "', " _
		&   "shippingLocState	= '" & validSQL(getStateDesc(shippingLocCountry,shippingLocState,shippingLocState2),"A") & "', " _
		&   "shippingLocCountry	= '" & validSQL(getCountryDesc(shippingLocCountry),"A") & "', " _
		&   "taxExempt			= '" & validSQL(taxExempt,"A") & "' " _
		&   "WHERE idOrder		=  " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)

end sub
'**********************************************************************
'Cancel Order
'**********************************************************************
sub CancelOrder()

	'Set CursorLocation of the Connection Object to Client
	connTemp.CursorLocation = adUseClient

	'BEGIN Transaction
	connTemp.BeginTrans
	
	'Delete cartRowsOptions
	mySQL = "DELETE FROM cartRowsOptions WHERE idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)

	'Delete cartRows
	mySQL = "DELETE FROM cartRows WHERE idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)
	 
	'Delete cartHead
	mySQL = "DELETE FROM cartHead WHERE idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)
	
	'END Transaction
	connTemp.CommitTrans
	
	'Set CursorLocation of the Connection Object back to Server
	connTemp.CursorLocation = adUseServer

end sub
'**********************************************************************
'Submit Order
'**********************************************************************
sub SubmitOrder()

	'Retrieve information we already have in cartHead so that we
	'can send emails, update order status, update discount code.
	mySQL="SELECT Total,Name,LastName,Email,paymentType," _
		& "       cardNumber,generalComments,deliverydate,discCode " _
		& "FROM   cartHead " _
		& "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	if not rstemp.eof then
		Total			= rstemp("Total")
		Name			= rstemp("name")
		LastName		= rstemp("LastName")
		Email			= rstemp("Email")
		paymentType		= rstemp("paymentType")
		cardNumber		= EnDeCrypt(Hex2Ascii(rstemp("cardNumber")),rc4Key)
		generalComments = rstemp("generalComments")
		deliverydate    = rstemp("deliverydate")
		discCode		= rstemp("discCode")
	end if
	call closeRS(rsTemp)
	
	'Update discount code to "Used" if it's a "Once Only" discount
	mySQL = "UPDATE discOrder " _
	      & "SET    discStatus = 'U' " _
		  & "WHERE  discOnceOnly = 'Y' " _
		  & "AND    discCode = '" & validSQL(discCode,"A") & "'"
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)
	
	'Get Order Date and determine Order Status
	orderDate = now()
	if Total = 0 then
		orderStatus = "1" 'Paid
	else
		orderStatus = "0" 'Pending
	end if
	
	'Update Order Status (and Stock if necessary)
	call updOrderStatus(idOrder,orderStatus,"N","Y","")
	
	'Update Order Dates
	mySQL = "UPDATE cartHead " _
	      & "SET    orderDate = '"	  & validSQL(orderDate,"A") & "', " _
	      & "       orderDateInt = '" & validSQL(dateInt(orderDate),"A") & "' " _
	      & "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)
	
	'Build Email Body
	customerEmail = ""
	mySQL = "SELECT configValLong " _
		&   "FROM   storeAdmin " _
		&   "WHERE  configVar = 'saveOrderEmail' " _
		&   "AND    adminType = 'T'"
	set rsTemp = openRSexecute(mySQL)
	if not rstemp.eof then
		customerEmail = trim(rsTemp("configValLong"))
	end if
	call closeRS(rsTemp)
	
	'Build Item detail to be sent with email (if necessary)
	if inStr(customerEmail,"#ITEMS#") > 0 then
	
		'Cart Rows
		mySQL = "SELECT idCartRow,quantity,unitPrice," _
			  & "       description,discAmt " _ 
			  & "FROM   cartRows " _ 
			  & "WHERE  idOrder = " & validSQL(idOrder,"I") 
		set rsTemp = openRSexecute(mySQL) 
		do while not rsTemp.EOF
			idCartRow			= rsTemp("idCartRow")
			quantity			= rsTemp("quantity")
			unitPrice			= rsTemp("unitPrice")
			description			= rsTemp("description")
			discAmt				= rsTemp("discAmt")
			optionGroupsTotal	= 0
			optionGroupsDesc	= ""
			
			'Cart Options
			mySQL = "SELECT optionPrice,optionDescrip " _
				  & "FROM   cartRowsOptions " _
				  & "WHERE  idCartRow = " & validSQL(idCartRow,"I")
			set rsTemp2 = openRSexecute(mySQL)
			do while not rstemp2.eof
				optionPrice			= rsTemp2("optionPrice")
				optionDescrip		= rsTemp2("optionDescrip")
				optionGroupsTotal	= optionGroupsTotal + optionPrice
				optionGroupsDesc	= optionGroupsDesc & "      " & optionDescrip & " (" & pCurrencySign & moneyS(optionPrice) & ")" & vbCrlf
				rstemp2.movenext
			loop
			call closeRS(rsTemp2)
			
			'Build items string
			customerEmailItems = customerEmailItems _
				& quantity & " x " _
				& description & "  @  " _
				& pCurrencySign & moneyS(Cdbl(optionGroupsTotal + unitPrice - discAmt)) & vbCrlf _
				& optionGroupsDesc
			
			'Add extra line feed if not at end of fle
			rsTemp.MoveNext
			if not rsTemp.EOF then
				customerEmailItems = customerEmailItems & vbCrlf
			end if
			
		loop 
		call CloseRS(rsTemp) 
		
	end if
	
	'Check for tags and replace
	customerEmail = replace(customerEmail,"#NAME#",name & " " & lastname)
	customerEmail = replace(customerEmail,"#DATE#",formatTheDate(date()))
	customerEmail = replace(customerEmail,"#ORDER#",pOrderPrefix & "-" & idOrder)
	customerEmail = replace(customerEmail,"#TOTAL#",pCurrencySign & moneyS(Total))
	customerEmail = replace(customerEmail,"#ITEMS#",customerEmailItems)
	customerEmail = replace(customerEmail,"#COMMT#",generalComments)
	customerEmail = replace(customerEmail,"#DELDATE#",deliverydate)
	customerEmail = replace(customerEmail,"#PAYMT#",paymentMsg(paymentType, total, cardNumber))
	customerEmail = replace(customerEmail,"#STORE#",pCompany)
	customerEmail = replace(customerEmail,"#SALES#",pEmailSales)

	'Send Email to Customer
	call sendmail (pCompany, pEmailSales, Email, pCompany & " " & langGenOrderNumber & " " & pOrderPrefix & "-" & idOrder, customerEmail, 0)

	'Send Email to Store Owner
	call sendmail (pCompany, pEmailAdmin, pEmailSales, langGenOrderNumber & " " & pOrderPrefix & "-" & idOrder, customerEmail, 0)

end sub
%>
