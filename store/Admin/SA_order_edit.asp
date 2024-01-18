<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Order Maintenance
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
<!--#include file="../Scripts/_INCrc4_.asp"-->
<%

'Database
dim mySQL, cn, rs, rs2

'cartHead
dim idOrder
dim idCust
dim orderDate
dim orderDateInt
dim randomKey
dim subTotal
dim taxTotal
dim shipmentTotal
dim handlingFeeTotal
dim Total
dim shipmentMethod
dim name
dim lastName
dim customerCompany
dim phone
dim email
dim address
dim city
dim locState
dim locCountry
dim zip
dim shippingName
dim shippingLastName
dim shippingPhone
dim shippingAddress
dim shippingCity
dim shippingLocState
dim shippingLocCountry
dim shippingZip
dim paymentType
dim cardType
dim cardNumber
dim cardExpMonth
dim cardExpYear
dim cardVerify
dim cardName
dim generalComments
dim deliverydate
dim orderStatus
dim auditInfo
dim storeComments
dim storeCommentsPriv
dim adjustAmount
dim adjustReason
dim discCode
dim discPerc
dim discTotal

'cartRows
dim idCartRow
dim idProduct
dim sku
dim quantity
dim unitPrice
dim unitWeight
dim description
dim downloadCount
dim downloadDate
dim discAmt

'CartRowsOptions
dim idCartRowOption
dim idOption
dim optionPrice
dim optionDescrip

'DiscProd
dim idDiscProd
dim discFromQty
dim discToQty

'Work Fields
dim action

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if  action <> "edit" _
and action <> "del"  _
and action <> "view" _
and action <> "inv" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'If this is an invoice, show different headers
if action = "inv" then
%>
	<html>
	<head>
		<title>Invoice</title>
		<style type="text/css">
		<!--
		BODY, B, TD, P {COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt;}
		-->
		</style>
	</head>
	<body>
<%
else
%>
	<!--#include file="_INCheader_.asp"-->
	<P align=left>
		<b><font size=3>Order Maintenance</font></b>
		<br><br>
	</P>
<%
end if

'Get idOrder
idOrder = trim(Request.QueryString("recId"))
if len(idOrder) = 0 then
	idOrder = trim(Request.Form("recId"))
end if
if idOrder = "" or not isNumeric(idOrder) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
end if

'Get cartHead Record
mySQL	= "SELECT idCust,orderDate,subTotal,taxTotal,shipmentTotal," _
		& "       Total,shipmentMethod,name,lastName,customerCompany," _
		& "       phone,email,address,city,locState,locCountry,zip," _
		& "       shippingName,shippingLastName,shippingPhone," _
		& "       shippingAddress,shippingCity,shippingLocState," _
		& "       shippingLocCountry,shippingZip,paymentType,cardType," _
		& "       cardNumber,cardExpMonth,cardExpYear,cardVerify," _
		& "       cardName,generalComments,deliverydate,orderStatus,auditInfo," _
		& "       adjustAmount,adjustReason,discCode,discPerc,discTotal," _
		& "       handlingFeeTotal,storeComments,storeCommentsPriv " _
		& "FROM   cartHead " _
		& "WHERE  idOrder=" & idOrder
set rs = openRSexecute(mySQL)
if rs.eof then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
else
	'NOTE : storeComments is assigned before storeCommentsPriv because 
	'we kept on getting the error 'Multiple-step OLE DB operation 
	'generated errors' on SQL Server 7.0 (Access was fine) if we 
	'attempted to assign storeCommentsPriv before storeComments. This 
	'is interesting because the general rule is that you should list 
	'TEXT/MEMO fields at the end of the SELECT, and read them first to 
	'last.
	storeComments		= rs("storeComments")
	storeCommentsPriv	= rs("storeCommentsPriv")
	idCust  			= rs("idCust")
	orderDate			= rs("orderDate")
	subTotal			= rs("subTotal")
	taxTotal			= rs("taxTotal")
	shipmentTotal		= rs("shipmentTotal")
	Total				= rs("Total")
	shipmentMethod		= trim(rs("shipmentMethod"))
	Name				= trim(rs("name"))
	LastName			= trim(rs("LastName"))
	CustomerCompany		= trim(rs("CustomerCompany"))
	Phone				= trim(rs("Phone"))
	Email				= trim(rs("Email"))
	Address				= trim(rs("Address"))
	City				= trim(rs("City"))
	locState			= trim(rs("locState"))
	locCountry			= trim(rs("locCountry"))
	Zip					= trim(rs("Zip"))
	shippingName		= trim(rs("shippingName"))
	shippingLastName	= trim(rs("shippingLastName"))
	shippingPhone		= trim(rs("shippingPhone"))
	shippingAddress		= trim(rs("shippingAddress"))
	ShippingCity		= trim(rs("ShippingCity"))
	shippingLocState	= trim(rs("shippingLocState"))
	shippingLocCountry	= trim(rs("shippingLocCountry"))
	shippingZip			= trim(rs("shippingZip"))
	paymentType			= trim(rs("paymentType"))
	cardType			= trim(rs("cardType"))
	cardNumber			= trim(EnDeCrypt(Hex2Ascii(rs("cardNumber")),rc4Key))
	cardExpMonth		= trim(rs("cardExpMonth"))
	cardExpYear			= trim(rs("cardExpYear"))
	cardVerify			= trim(rs("cardVerify"))
	cardName			= trim(rs("cardName"))
	generalComments		= trim(rs("generalComments"))
	deliverydate		= trim(rs("deliverydate"))
	orderStatus			= rs("orderStatus")
	auditInfo			= rs("auditInfo")
	adjustAmount		= rs("adjustAmount")
	adjustReason		= trim(rs("adjustReason"))
	discCode			= trim(rs("discCode"))
	discPerc			= trim(rs("discPerc"))
	discTotal			= trim(rs("discTotal"))
	handlingFeeTotal	= rs("handlingFeeTotal")
	
	'Cater for orders entered before order discounts were added
	if isNull(discPerc) then
		discPerc = 0.00
	end if
	if isNull(discTotal) then
		discTotal = 0.00
	end if
	
end if
call closeRS(rs)

'Edit
if action = "edit" then
	if len(trim(Request.QueryString("msg"))) > 0 then
%>
		<font color=red><%=Request.QueryString("msg")%></font><br><br>
<%
	end if
%>
	<span class="textBlockHead">Edit Order</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<!-- Start Table Outer Cell -->
	<table border=0 cellspacing=0 cellpadding=5 width="450" class="textBlock">
	<form method="post" action="SA_order_exec.asp" name="form1">
	<TR><TD>
	
	<!-- General Info & Addresses -->
	<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="1" WIDTH="100%">
	<TR>
		<TD align=left>
			<table border=0>
				<tr>
					<td valign=top nowrap><b>Order Number</b>&nbsp;</td>
					<td valign=top nowrap>
						<%=pOrderPrefix & "-" & idOrder%>&nbsp;&nbsp;&nbsp;&nbsp;
						[ 
						<a href="SA_cust_edit.asp?action=edit&recID=<%=idCust%>">Edit Customer</a> | 
						<a href="email.asp?emailTo=<%=server.URLEncode(email & " ")%>&emailToName=<%=server.URLEncode(name & " " & LastName)%>&emailSubj=<%=server.URLEncode(pCompany & " - Order " & pOrderPrefix & "-" & idOrder)%>">Send Email</a> 
						]
					</td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Order Date</b>&nbsp;</td>
					<td valign=top><%=formatTheDate(orderDate)%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Order Status</b>&nbsp;</td>
					<td valign=top>
						<select name=orderStatus id=orderStatus size=1>
							<option value="U" <%=checkMatch(orderStatus,"U")%>>Unfinalized</option>
							<option value="S" <%=checkMatch(orderStatus,"S")%>>Saved</option>
							<option value="0" <%=checkMatch(orderStatus,"0")%>>Pending</option>
							<option value="1" <%=checkMatch(orderStatus,"1")%>>Paid</option>
							<option value="2" <%=checkMatch(orderStatus,"2")%>>Shipped</option>
							<option value="7" <%=checkMatch(orderStatus,"7")%>>Complete</option>
							<option value="9" <%=checkMatch(orderStatus,"9")%>>Cancelled</option>
						</select><br>
						<input type=checkbox name=orderStatusMail     value="Y" checked> Email Customer if status changes.<br>
						<input type=checkbox name=orderStatusStockAdj value="Y" checked> Adjust Stock Levels if status changes.<br>
					</td>
				</tr>
				<tr>
					<td valign=top>
						<b>Store&nbsp;Comments</b>&nbsp;<br>
						(Can be viewed by the Customer)
					</td>
					<td valign=top>
						<textarea name=storeComments cols=35 rows=6><%=storeComments%></textarea>
					</td>
				</tr>
				<tr>
					<td valign=top>
						<b>Private&nbsp;Comments</b>&nbsp;<br>
						(Can NOT be viewed by the Customer)
					</td>
					<td valign=top>
						<textarea name=storeCommentsPriv cols=35 rows=6><%=storeCommentsPriv%></textarea>
					</td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Email</b>&nbsp;</td>
					<td valign=top>
						<input type=text name=email id=email size=30 maxlength=100 value="<%=email%>">
					</td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Customer ID</b>&nbsp;</td>
					<td valign=top nowrap><%=idCust%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Audit Info</b>&nbsp;</td>
					<td valign=top><%=auditInfo%></td>
				</tr>
			</table>
		</TD>
    </TR>
	<tr> 
		<td align=left>
			<table border="0" cellspacing="1" cellpadding="3" width="100%">
				<tr>
					<td bgcolor="#DDDDDD">&nbsp;</td>
					<td bgcolor="#DDDDDD"><b>Billing</b></td>
					<td bgcolor="#DDDDDD"><b>Shipping</b></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>First Name</b>&nbsp;</td>
					<td>
						<input type=text name=Name id=Name size=20 maxlength=50 value="<%=Name%>">
					</td>
					<td>
						<input type=text name=shippingName id=shippingName size=20 maxlength=50 value="<%=shippingName%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Last Name</b>&nbsp;</td>
					<td>
						<input type=text name=LastName id=LastName size=20 maxlength=50 value="<%=LastName%>">
					</td>
					<td>
						<input type=text name=shippingLastName id=shippingLastName size=20 maxlength=50 value="<%=shippingLastName%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Address</b>&nbsp;</td>
					<td>
						<input type=text name=address id=address size=20 maxlength=70 value="<%=address%>">
					</td>
					<td>
						<input type=text name=shippingAddress id=shippingAddress size=20 maxlength=70 value="<%=shippingAddress%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>City</b>&nbsp;</td>
					<td>
						<input type=text name=city id=city size=20 maxlength=50 value="<%=city%>">
					</td>
					<td>
						<input type=text name=shippingCity id=shippingCity size=20 maxlength=50 value="<%=shippingCity%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Zip/PCode</b>&nbsp;</td>
					<td>
						<input type=text name=zip id=zip size=10 maxlength=10 value="<%=zip%>">
					</td>
					<td>
						<input type=text name=shippingZip id=shippingZip size=10 maxlength=10 value="<%=shippingZip%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>State/Province</b>&nbsp;</td>
					<td>
						<input type=text name=locState id=locState size=20 maxlength=100 value="<%=locState%>">
					</td>
					<td>
						<input type=text name=shippingLocState id=shippingLocState size=20 maxlength=100 value="<%=shippingLocState%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Country</b>&nbsp;</td>
					<td>
						<input type=text name=locCountry id=locCountry size=20 maxlength=100 value="<%=locCountry%>">
					</td>
					<td>
						<input type=text name=shippingLocCountry id=shippingLocCountry size=20 maxlength=100 value="<%=shippingLocCountry%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Phone</b>&nbsp;</td>
					<td>
						<input type=text name=phone id=phone size=20 maxlength=30 value="<%=phone%>">
					</td>
					<td>
						<input type=text name=shippingPhone id=shippingPhone size=20 maxlength=30 value="<%=shippingPhone%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Company</b>&nbsp;</td>
					<td colspan=2>
						<input type=text name=customerCompany id=customerCompany size=20 maxlength=50 value="<%=customerCompany%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Shipping</b>&nbsp;</td>
					<td colspan=2>
						<input type=text name=shipmentMethod id=shipmentMethod size=40 maxlength=100 value="<%=shipmentMethod%>">
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Payment</b>&nbsp;</td>
					<td colspan=2>
						<select name="paymentType" size=1>
							<option value="">------- Select -------</option>
							<option value="MailIn"       <%if lCase(paymentType)="mailin"		then Response.Write "selected" end if%>><%=paymentMsg("mailin",1,"")%>
							<option value="CallIn"       <%if lCase(paymentType)="callin"		then Response.Write "selected" end if%>><%=paymentMsg("callin",1,"")%>
							<option value="FaxIn"        <%if lCase(paymentType)="faxin"		then Response.Write "selected" end if%>><%=paymentMsg("faxin",1,"")%>
							<option value="COD"          <%if lCase(paymentType)="cod"			then Response.Write "selected" end if%>><%=paymentMsg("cod",1,"")%>
							<option value="CreditCard"   <%if lCase(paymentType)="creditcard"	then Response.Write "selected" end if%>><%=paymentMsg("creditcard",1,"")%>
							<option value="PayPal"       <%if lCase(paymentType)="paypal"		then Response.Write "selected" end if%>><%=paymentMsg("paypal",1,"")%>
							<option value="2CheckOut"    <%if lCase(paymentType)="2checkout"	then Response.Write "selected" end if%>><%=paymentMsg("2checkout",1,"")%>
							<option value="AuthorizeNet" <%if lCase(paymentType)="authorizenet"	then Response.Write "selected" end if%>><%=paymentMsg("authorizenet",1,"")%>
							<option value="Custom"		 <%if lCase(paymentType)="custom"		then Response.Write "selected" end if%>><%=paymentMsg("custom",1,"")%>
						</select>
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Comments</b>&nbsp;</td>
					<td colspan=2><i><%=emptyString(generalComments,"None")%></i></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Delivery date</b>&nbsp;</td>
					<td colspan=2><i><%=emptyString(Deliverydate,"None")%></i></td>
				</tr>
<%
				if len(cardNumber) > 0 then
%>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Type</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardType%></font></td>
					</tr>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Number</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardNumber%></font></td>
					</tr>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Expire</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardExpMonth & "/" & cardExpYear%>&nbsp;&nbsp;(MM/YYYY)</font></td>
					</tr>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Verif. #</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardVerify%></font></td>
					</tr>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Name</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardName%></font></td>
					</tr>
<%
				end if
%>
			</table>
		</td>
	</tr>
	</TABLE>
	
	<!-- Items & Totals -->
	<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="1" WIDTH="100%">
<%
	'Display all the items and options for this order
	call showOrderItems(idOrder)
%>
	<tr> 
		<td colspan="2" align=right bgcolor="#DDDDDD"><b>Sub Total:&nbsp;&nbsp;</b></td>
		<td align=left bgcolor="#DDDDDD" nowrap><b><%=pCurrencySign & moneyD(subTotal)%></b></td>
	</tr>
	<tr> 
		<td colspan="2" align=right>
<%
			if discTotal > 0 then
				Response.Write "<i>" & discCode & " (" & formatNumber(discPerc,2) & "%)" & "</i> - "
			end if
%>
			<b>Discount Code:&nbsp;&nbsp;</b>
		</td>
		<td align=left nowrap> 
<%
			Response.Write pCurrencySign & moneyD(discTotal)
			if discTotal > 0 then
				Response.Write "&nbsp;&nbsp;(-)"
			end if
%>
		</td>
	</tr>
	<tr> 
		<td colspan="2" align=right><b>Sub Total:&nbsp;&nbsp;</b></td>
		<td align=left nowrap><b><%=pCurrencySign & moneyD(subTotal - discTotal)%></b></td>
	</tr>
	<tr> 
		<td colspan="2" align=right><b>Shipping:&nbsp;&nbsp;</b></td>
		<td align=left nowrap><%=pCurrencySign & moneyD(shipmentTotal)%></td>
	</tr>
	<tr> 
		<td colspan="2" align=right><b>Handling Fee:&nbsp;&nbsp;</b></td>
		<td align=left nowrap><%=pCurrencySign & moneyD(handlingFeeTotal)%></td>
	</tr>
	<tr> 
		<td colspan="2" align=right><b>Tax:&nbsp;&nbsp;</b></td>     
		<td align=left nowrap><%=pCurrencySign & moneyD(taxTotal)%>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" align=right bgcolor="#DDDDDD"><b>Total:&nbsp;&nbsp;</b></td>
		<td align=left valign=top bgcolor="#DDDDDD" nowrap><b><%=pCurrencySign & moneyD(Total)%></b></td>
	</tr>
	<tr>
		<td colspan="3" align=left>
			<i>Adjustment Reason and Amount. See HELP for more info.</i>
		</td>
	</tr>
	<tr> 
		<td colspan="2" align=left>
			<input type=text name=adjustReason id=adjustReason size=50 maxlength=250 value="<%=adjustReason%>">
		</td>     
		<td align=left nowrap>
			<%=replace(pCurrencySign," ","&nbsp;")%><input type=text name=adjustAmount id=adjustAmount size=5 maxlength=10 value="<%=moneyD(adjustAmount)%>">
		</td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="3">
			<input type=hidden name=idOrder         id=idOrder         value="<%=idOrder%>">
			<input type=hidden name=total           id=total           value="<%=total%>">
			<input type=hidden name=totalNoAdjust   id=totalNoAdjust   value="<%=subTotal - discTotal + shipmentTotal + handlingFeeTotal + taxTotal%>">
			<input type=hidden name=action          id=action          value="edit">
			<input type=submit name=submit1         id=submit1         value="Update Order">
		</td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
	</TABLE>
	
	<!-- End Table outer Cell -->
	</TD></TR>
	</FORM>
	</TABLE>
<%
end if

'View / Invoice / Delete
if action = "view" or action = "inv" or action = "del" then

	'Delete
	if action = "del" then
%>
		<span class="textBlockHead">Delete Order</span>
		&nbsp;<%call maintNavLinks()%><br>

		<form method="post" action="SA_order_exec.asp" name="form4">
			<font color=red>Are you sure you want to Delete this Order?</font>
			<input type=hidden name=idOrder id=idOrder value="<%=idOrder%>">
			<input type=hidden name=action  id=action  value="del">
			<input type=submit name=submit1 id=submit1 value="Yes">
		</form>
<%
	'View
	elseif action = "view" then
%>
		<span class="textBlockHead">View Order</span>
		&nbsp;<%call maintNavLinks()%><br><br>
<%
	'Invoice
	elseif action = "inv" then
%>
		<table border="0" cellpadding="3" cellspacing="0" width="450">
		<tr>
			<td nowrap bgColor="#DDDDDD">
				<b><font size=3><%=pCompany%></font></b>
			</td>
			<td align=right nowrap bgColor="#DDDDDD">
				<b><font size=3>Invoice</font></b>
			</td>
		</tr>
		</table>
<%
	end if
%>
	<!-- Start Table Outer Cell -->
	<table border=0 cellspacing=0 cellpadding=5 width="450" class="textBlock">
	<TR><TD>
	
	<!-- General Info & Addresses -->
	<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="1" WIDTH="100%">
	<TR>
		<TD align=left>
			<table border=0 width="100%">
<%
				'If this is an invoice, we adjust the next couple of 
				'rows a little bit to include the store address
				if action = "inv" then
%>
				<tr>
					<td valign=top nowrap width="50%">
						<%=replace(pCompanyAddr,chr(10),"<br>")%>
					</td>
					<td valign=top nowrap width="50%">
						<table border=0 cellspacing=0 cellpadding=0>
						<tr>	
							<td nowrap><b>Order Number</b>&nbsp;</td>
							<td nowrap><%=pOrderPrefix & "-" & idOrder%></td>
						</tr>
						<tr>
							<td nowrap><b>Order Date</b>&nbsp;</td>
							<td nowrap><%=formatTheDate(orderDate)%></td>
						</tr>
						<tr>
							<td nowrap><b>Order Status</b>&nbsp;</td>
							<td nowrap><%=orderStatusDesc(orderStatus)%></td>
						</tr>
						</table>
					</td>
				</tr>
<%
				else
%>
				<tr>
					<td valign=top nowrap><b>Order Number</b>&nbsp;</td>
					<td valign=top nowrap>
						<%=pOrderPrefix & "-" & idOrder%>&nbsp;&nbsp;&nbsp;&nbsp;
						[ 
						<a href="SA_cust_edit.asp?action=edit&recID=<%=idCust%>">Edit Customer</a> | 
						<a href="email.asp?emailTo=<%=server.URLEncode(email & " ")%>&emailToName=<%=server.URLEncode(name & " " & LastName)%>&emailSubj=<%=server.URLEncode(pCompany & " - Order " & pOrderPrefix & "-" & idOrder)%>">Send Email</a> 
						]
					</td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Order Date</b>&nbsp;</td>
					<td valign=top><%=formatTheDate(orderDate)%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Order Status</b>&nbsp;</td>
					<td valign=top><%=orderStatusDesc(orderStatus)%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Store Comments</b>&nbsp;</td>
					<td valign=top><%=replace(emptyString(storeComments,"None"),Chr(10),"<br>")%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Private Comments</b>&nbsp;</td>
					<td valign=top><font color=red><%=replace(emptyString(storeCommentsPriv,"None"),Chr(10),"<br>")%></font></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Email</b>&nbsp;</td>
					<td valign=top><font color=red><%=email%></font></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Customer ID</b>&nbsp;</td>
					<td valign=top nowrap><%=idCust%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Audit Info</b>&nbsp;</td>
					<td valign=top><font color=red><%=auditInfo%></font></td>
				</tr>
<%
				end if
%>
			</table>
		</TD>
    </TR>
	<tr> 
		<td align=left>
			<table border="0" cellspacing="1" cellpadding="3" width="100%">
				<tr>
					<td bgcolor="#DDDDDD">&nbsp;</td>
					<td width="50%" bgcolor="#DDDDDD"><b>Billing</b></td>
					<td width="50%" bgcolor="#DDDDDD"><b>Shipping</b></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Name</b>&nbsp;</td>
					<td><%=Name & " " & LastName%></td>
					<td><%=emptyString(shippingName,Name) & " " & emptyString(shippingLastName,LastName)%></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Address</b>&nbsp;</td>
					<td><%=address%></td>
					<td><%=emptyString(shippingAddress,address)%></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>City</b>&nbsp;</td>
					<td><%=city%></td>
					<td><%=emptyString(shippingCity,city)%></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Zip/PCode</b>&nbsp;</td>
					<td><%=zip%></td>
					<td><%=emptyString(shippingZip,zip)%></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Location</b>&nbsp;</td>
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
					<td nowrap bgcolor="#DDDDDD"><b>Phone</b>&nbsp;</td>
					<td><%=phone%></td>
					<td><%=emptyString(shippingPhone,phone)%></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Company</b>&nbsp;</td>
					<td><%=customerCompany%></td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Shipping</b>&nbsp;</td>
					<td colspan=2><i><%=shipmentMethod%></i></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Payment</b>&nbsp;</td>
					<td colspan=2>
						<i><%=paymentMsg(paymentType, total, cardNumber)%></i>
					</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Comments</b>&nbsp;</td>
					<td colspan=2><i><%=emptyString(generalComments,"None")%></i></td>
				</tr>
				<tr>
					<td nowrap bgcolor="#DDDDDD"><b>Delivery date</b>&nbsp;</td>
					<td colspan=2><i><%=emptyString(deliverydate,"None")%></i></td>
				</tr>
<%
				'Show credit card info if this is a Credit Card payment 
				'and not an invoice.
				if len(cardNumber) > 0 and action <> "inv" then
%>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Type</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardType%></font></td>
					</tr>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Number</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardNumber%></font></td>
					</tr>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Expire</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardExpMonth & "/" & cardExpYear%>&nbsp;&nbsp;(MM/YYYY)</font></td>
					</tr>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Verif. #</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardVerify%></font></td>
					</tr>
					<tr>
						<td nowrap bgcolor="#DDDDDD"><b>Card Name</b>&nbsp;</td>
						<td colspan=2><font color=red><%=cardName%></font></td>
					</tr>
<%
				end if
%>
			</table>
		</td>
	</tr>
	</TABLE>
	
	<!-- Items & Totals -->
	<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="1" WIDTH="100%">
<%
	'Display all the items and options for this order
	call showOrderItems(idOrder)
%>
	<tr> 
		<td colspan="2" align=right bgcolor="#DDDDDD"><b>Sub Total:&nbsp;&nbsp;</b></td>
		<td align=left bgcolor="#DDDDDD" nowrap><b><%=pCurrencySign & moneyD(subTotal)%></b></td>
	</tr>
	<tr> 
		<td colspan="2" align=right>
<%
			if discTotal > 0 then
				Response.Write "<i>" & discCode & " (" & formatNumber(discPerc,2) & "%)" & "</i> - "
			end if
%>
			<b>Discount Code:&nbsp;&nbsp;</b>
		</td>
		<td align=left nowrap> 
<%
			Response.Write pCurrencySign & moneyD(discTotal)
			if discTotal > 0 then
				Response.Write "&nbsp;&nbsp;(-)"
			end if
%>
		</td>
	</tr>
	<tr> 
		<td colspan="2" align=right><b>Sub Total:&nbsp;&nbsp;</b></td>
		<td align=left nowrap><b><%=pCurrencySign & moneyD(subTotal - discTotal)%></b></td>
	</tr>
	<tr> 
		<td colspan="2" align=right><b>Shipping:&nbsp;&nbsp;</b></td>
		<td align=left nowrap><%=pCurrencySign & moneyD(shipmentTotal)%></td>
	</tr>
	<tr> 
		<td colspan="2" align=right><b>Handling Fee:&nbsp;&nbsp;</b></td>
		<td align=left nowrap><%=pCurrencySign & moneyD(handlingFeeTotal)%></td>
	</tr>
	<tr> 
		<td colspan="2" align=right><b>Tax:&nbsp;&nbsp;</b></td>     
		<td align=left nowrap><%=pCurrencySign & moneyD(taxTotal)%>&nbsp;</td>
	</tr>
	<tr> 
		<td colspan="2" align=right>*<b>Adjustment:&nbsp;&nbsp;</b></td>     
		<td align=left nowrap>
<%
			if isNumeric(adjustAmount) then
				Response.Write pCurrencySign & moneyD(adjustAmount)
			else
				Response.Write pCurrencySign & moneyD("0")
			end if
%>
		</td>
	</tr>
	<tr>
		<td colspan="2" align=right bgcolor="#DDDDDD"><b>Total:&nbsp;&nbsp;</b></td>
		<td align=left valign=top bgcolor="#DDDDDD" nowrap><b><%=pCurrencySign & moneyD(Total)%></b></td>
	</tr>
	<tr>
		<td colspan="3">
			*<b>Adjustment :</b> 
<%
			if len(adjustReason) > 0 then
				Response.Write adjustReason
			else
				Response.Write "No Adjustment(s) for this Order&nbsp;&nbsp;"
			end if
%>
		</td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
<%
	'Only show message if this is not an invoice
	if action <> "inv" then
%>
	<TR> 
		<TD COLSPAN="3">
			NOTE : This Order as it is displayed here closely 
			resembles the Order as it is displayed to the Customer 
			when they log on to their Account and view an Order's 
			detail. However, <font color=red>fields shown in 
			red</font> are only viewable by the store Administrator.
		</TD>
    </TR>
<%
	end if
%>
	</TABLE>
	
	<!-- End Table outer Cell -->
	</TD></TR>
	</TABLE>

<%
end if

if action = "edit" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>
	
	<b>Order Status</b> - Changing the Order Status will 
	automatically result in the following actions also taking place :<br><br>
	
	1. Send an email to the Customer notifying them of the change (if 
	"Email Customer if status changes" box is checked).<br><br>
	
	2. Update the Stock Levels of all Products involved (if 
	"Adjust Stock Levels if status changes" box is checked).<br><br>
	
	3. Append the Date/Time to the "Store Comments" field.<br><br>
	
	<b>Adjustments</b> - Optionally adjust the Order Total by 
	entering an adjustment Reason and Amount in the space 
	provided. A negative value will decrease, and a positive 
	value will increase the ORIGINAL Order Total by the amount 
	that you entered. The customer will be able to see the 
	Adjustment Reason and Amount when they view the order via 
	their Account. To reset the Order Total to the original 
	amount, simply set the Adjustment Amount to "0.00".<br><br>
	
	<b>Store Comments</b> - Any text entered here will be 
	viewable by the Customer. It's usefull if for some reason there is 
	a problem with the Order and you want to communicate this to your 
	Customer.<br><br>
	
	<b>Private Comments</b> - Text entered here will NOT be 
	viewable by the Customer. Use this field to store information that 
	is confidential to your Store.<br><br>
	
	<b>Other</b> - Several other fields are modifiable. These fields 
	are typically entered by the Customer when placing the Order. Under 
	certain circumstances it may be necessary to change this information, 
	mainly due to user error when the Order is placed. Once an Order is 
	placed, the Customer can no longer make any changes. The modifications 
	made to these fields are not strictly checked. This is to allow you 
	maximum flexibility should you need to make changes.<br><br>
	
	</td></tr>
	</table>
<%
end if

'Close Database Connection
call closedb()

'If this is an invoice, show different footers
if action = "inv" then
%>
	</body></html>
<%
else
%>
	<!--#include file="_INCfooter_.asp"-->
<%
end if

'*********************************************************************
'Display the Order's Items
'*********************************************************************
sub showOrderItems(idOrder)

	'Declare local vars
	dim optionGroupsTotal
%>
	<tr> 
		<td width="10%" bgcolor="#DDDDDD" nowrap><b>Qty</b></td>
		<td width="80%" bgcolor="#DDDDDD" nowrap><b>Item Description</b></td>
		<td width="10%" bgcolor="#DDDDDD" nowrap><b>Sub Total</b></td>
	</tr>
<%
	'Get all rows for this cart
	mySQL="SELECT idCartRow,idProduct,quantity,unitPrice," _
		& "       description,sku,downloadCount,downloadDate," _
		& "       discAmt " _
	    & "FROM   cartRows " _
	    & "WHERE  cartRows.idOrder=" & idOrder & " " _
	    & "ORDER BY description"
	set rs = openRSexecute(mySQL)
	do while not rs.eof
	
		'Assign record values to local values
		idCartRow		= rs("idCartRow")
		idProduct		= rs("idProduct")
		quantity		= rs("quantity")
		unitPrice		= rs("unitPrice")
		description		= rs("description")
		sku				= rs("sku")
		downloadCount	= rs("downloadCount")
		downloadDate	= rs("downloadDate")
		discAmt			= rs("discAmt")
		
		'Cater for orders entered before discounts were added
		if isNull(discAmt) then
			discAmt = 0.00
		end if
%> 
		<tr> 
			<td nowrap valign=top><%=quantity%></td>
			
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
				response.write " " & Description & " - <i>" & pCurrencySign & moneyD(unitPrice) & "</i><br>"
					
				'Write Discount (if any)
				if discAmt > 0 then
					Response.Write "* <i>Discount - " & pCurrencySign & moneyD(discAmt) & "</i><br>"
				end if
					
				'Get all options for this row
				optionGroupsTotal = 0
				mySQL = "SELECT optionPrice,optionDescrip " _
					  & "FROM   cartRowsOptions " _
					  & "WHERE  idCartRow=" & IDCartRow
				set rs2 = openRSexecute(mySQL)
				do while not rs2.eof
					
					'Assign record values to local values
					optionDescrip = rs2("optionDescrip")
					optionPrice	  = rs2("optionPrice")
						
					'Write cartRowOptions line(s) (options)
					Response.Write "* <i>" & optionDescrip
					if optionPrice <> 0 then
						Response.Write " - " & pCurrencySign & moneyD(optionPrice)
					end if
					Response.Write "</i><br>"

					'Calculate options Sub Total
					optionGroupsTotal = optionGroupsTotal + optionPrice        
						
					rs2.movenext
				loop
				call closeRS(rs2)
				
				'Display downloadCount and downloadDate (if not invoice)
				if isNumeric(downloadCount) and action <> "inv" then
					if downloadCount > 0 then
						Response.Write "<font color=red>Downloaded " & downloadCount & " times since '" & formatIntDate(downloadDate) & "'.</font>" 
					end if
				end if
%>
			</td>
			<td nowrap valign=top>
				<%=pCurrencySign & moneyD(Cdbl(quantity * (optionGroupsTotal + unitPrice - discAmt)))%> 
			</td>
		</tr>
<%
		rs.movenext
	loop
	call closeRS(rs)

end sub
'*********************************************************************
'Format the internal integer date
'*********************************************************************
function formatIntDate(str1)
	
	if len(trim(str1))=14 and isnumeric(str1) then
		formatIntDate = "" _
			& left(str1,4)  & "/" _
			& mid(str1,5,2) & "/" _
			& mid(str1,7,2) & " " _
			& mid(str1,9,2) & ":" _
			& mid(str1,11,2)
	else
		formatIntDate = str1
	end if

end function
'*********************************************************************
'Create Navigation Links
'*********************************************************************
sub maintNavLinks()
%>
	[ 
	<a href="SA_order.asp?recallCookie=1">List Orders</a> | 
	<a href="SA_order_edit.asp?action=view&recid=<%=idOrder%>">View</a> | 
	<a href="SA_order_edit.asp?action=edit&recid=<%=idOrder%>">Edit</a> | 
	<a href="SA_order_edit.asp?action=inv&recid=<%=idOrder%>" target="_blank">Invoice</a> | 
	<a href="SA_order_edit.asp?action=del&recid=<%=idOrder%>">Delete</a> 
<%	if pAuthNet = -1 then %>					
	| <a href="SA_authnet.asp?recid=<%=idOrder%>">Authorize</a> 
<%	end if %>
	]
<%
end sub
%>