<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Get shipping info, payment info
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
dim i, f				'Indexes
dim totalShipPrice		'Total price    - non-free shipping items
dim totalShipWeight		'Total weight   - non-free shipping items
dim totalShipItems		'Total quantity - non-free shipping items
dim shipParms			'Array of parameters passed to online routines
dim shipArray			'Array of Shipping Methods & Rates
redim shipArray(100,1)	'Redimension array to appropriate size
dim pCCTypeArr			'Array of Valid Credit Card Types
dim arrayErrors			'Array of errors on the form (if any)
dim shipDetails
dim paymentRequired
dim formID
dim shippingLocCountryDesc

'Customer
dim locState
dim locCountry
dim zip
dim shippingLocState
dim shippingLocCountry
dim shippingZip

'cartHead
dim subTotal
dim taxTotal
dim shipmentTotal
dim handlingFeeTotal
dim Total
dim shipmentMethod
dim paymentType
dim cardType
dim cardNumber
dim cardExpMonth
dim cardExpYear
dim cardName
dim cardVerify
dim generalComments
dim deliverydate

'shipRates
dim locShipZone
dim idShipMethod
dim unitsFrom
dim unitsTo
dim unitType
dim addAmt
dim addPerc

'shipMethod
dim shipDesc

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

'Double-check that the Customer is logged on
if isNull(idCust) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrNotLoggedIn)
end if

'Get Form ID
formID = trim(Request.Form("formID"))

'Before we display the form for the first time, do some checks
if formID = "" then

	'Retrieve all available fields from DB
	mySQL="SELECT a.shipmentMethod,a.shipmentTotal,a.cardType," _
		& "       a.cardNumber,a.cardExpMonth,a.cardExpYear,a.cardName," _
		& "       a.cardVerify,a.generalComments,a.deliverydate,b.paymentType " _
	    & "FROM   cartHead a, customer b " _
	    & "WHERE  a.idOrder = " & validSQL(idOrder,"I") & " " _
	    & "AND    b.idCust  = " & validSQL(idCust,"I")  & " "
	set rsTemp = openRSexecute(mySQL)
	if not rstemp.eof then
		shipmentMethod  = trim(rstemp("shipmentMethod")&"")
		shipmentTotal   = trim(rstemp("shipmentTotal")&"")
		shipDetails     = shipmentTotal & "|" & shipmentMethod
		paymentType     = trim(rstemp("paymentType")&"")
		cardType		= trim(rstemp("cardType")&"")
		cardNumber		= trim(EnDeCrypt(Hex2Ascii(rstemp("cardNumber")),rc4Key)&"")
		cardExpMonth	= trim(rstemp("cardExpMonth")&"")
		cardExpYear		= trim(rstemp("cardExpYear")&"")
		cardName		= trim(rstemp("cardName")&"")
		cardVerify		= trim(rstemp("cardVerify")&"")
		generalComments	= trim(rstemp("generalComments")&"")
		deliverydate	= trim(rstemp("deliverydate")&"")

	else
		'No cartHead Record on DB (which is highly unlikely because
		'cartHead record has already been tested in sessionCart()
		'at the beginning of this script).
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvOrder)
	end if
	call closeRS(rsTemp)
	
end if
	
'Check if the Customer clicked the "Next" button
if formID = "02" then

	'Get values from the calling form
	shipDetails = validHTML(request.form("shipDetails"))
	if inStr(shipDetails,"|") > 0 then
		shipmentTotal  = mid(shipDetails,1,instr(shipDetails,"|")-1)
		shipmentMethod = mid(shipDetails,instr(shipDetails,"|")+1)
	else
		shipmentTotal  = ""
		shipmentMethod = ""
	end if
	handlingFeeTotal = validHTML(request.form("handlingFeeTotal"))
	paymentRequired  = validHTML(request.form("paymentRequired"))
	paymentType		 = validHTML(request.form("paymentType"))
	cardType		 = validHTML(request.form("cardType"))
	cardNumber		 = validHTML(request.form("cardNumber"))
	cardExpMonth	 = validHTML(request.form("cardExpMonth"))
	cardExpYear		 = validHTML(request.form("cardExpYear"))
	cardName		 = validHTML(request.form("cardName"))
	cardVerify		 = validHTML(request.form("cardVerify"))
	generalComments	 = validHTML(request.form("generalComments"))
	deliverydate	 = validHTML(request.form("deliverydate"))

	
	
	
    'Validate Shipping
	if len(shipmentTotal) = 0 or not isNumeric(shipmentTotal) then
		arrayErrors = arrayErrors & "|shipDetails"
	end if
	if len(shipmentMethod) = 0 then
		arrayErrors = arrayErrors & "|shipDetails"
	end if
	
	'Validate Handling Fee
	if len(handlingFeeTotal) = 0 or not isNumeric(handlingFeeTotal) then
		arrayErrors = arrayErrors & "|handlingFeeTotal"
	end if

	'Validate Credit Card Info (if Required for this Order)
	if lCase(paymentType) = "creditcard" and paymentRequired = "Y" then
		'Card Type
		if len(cardType) = 0 then
			arrayErrors = arrayErrors & "|cardType"
		end if
		'Card Number
		if not isCreditCard(cardNumber) then
			arrayErrors = arrayErrors & "|cardNumber"
		end if
		'Card Month
		if isEmpty(cardExpMonth) or not isNumeric(cardExpMonth) then
			arrayErrors = arrayErrors & "|cardExpMonth"
		end if
		'Card Year
		if isEmpty(cardExpYear)or not isNumeric(cardExpYear) then
			arrayErrors = arrayErrors & "|cardExpYear"
		end if
		'Card Month + Year not expired
		if not isDate(cardExpMonth & "/01/" & cardExpYear) then
			arrayErrors = arrayErrors & "|cardExpMonth"
			arrayErrors = arrayErrors & "|cardExpYear"
		else
			if date() > CDate(cardExpMonth & "/01/" & cardExpYear) then
				arrayErrors = arrayErrors & "|cardExpMonth"
				arrayErrors = arrayErrors & "|cardExpYear"
			end if
		end if
		'Card Name
		if len(cardName) = 0 then
			arrayErrors = arrayErrors & "|cardName"
		end if
	end if
	
	'Validate Comments
	if len(generalComments) > 250 then
		arrayErrors = arrayErrors & "|generalComments"
	end if
	
	'Validate delivery date
	if len(deliverydate) = 0 then
		arrayErrors = arrayErrors & "|deliverydate"
	end if
	
	'There were no errors
	if len(trim(arrayErrors)) = 0 then
	
		'Update Shopping Cart on DB
		mySQL = "UPDATE cartHead SET " _
			&   "shipmentMethod	  = '" & validSQL(shipmentMethod,"A") & "', " _
			&   "shipmentTotal	  = "  & validSQL(shipmentTotal,"D") & ", " _
			&   "handlingFeeTotal = "  & validSQL(handlingFeeTotal,"D") & ", " _
			&   "paymentType	  = '" & validSQL(paymentType,"A") & "', " _
			&   "cardType		  = '" & validSQL(cardType,"A") & "', " _
			&   "cardNumber		  = '" & validSQL(Ascii2Hex(EnDeCrypt(cardNumber,rc4Key)),"A") & "', " _
			&   "cardExpMonth	  = '" & validSQL(cardExpMonth,"A") & "', " _
			&   "cardExpYear	  = '" & validSQL(cardExpYear,"A") & "', " _
			&   "cardName		  = '" & validSQL(cardName,"A") & "', " _
			&   "cardVerify		  = '" & validSQL(cardVerify,"A") & "', " _
			&   "generalComments  = '" & validSQL(generalComments,"A") & "', " _
			&   "deliverydate  = '" & validSQL(deliverydate,"A") & "' " _

			&   "WHERE idOrder	  =  " & validSQL(idOrder,"I")
		set rsTemp = openRSexecute(mySQL)
		call closeRS(rsTemp)
		
		'Onto next page
		Response.Redirect "40_SubmitOrder.asp"
	
	end if
	
end if

'If we get this far, it's either because this script was called by 
'another script, or the script called itself but failed some checks.

'*********************************************************************
'***********   S T A R T - S H I P P I N G   R A T E S   *************
'*********************************************************************

'-------------------------------------
' Prepare to calculate shipping rates
'-------------------------------------

'Get Shipping Location
mySQL = "SELECT locState,locCountry,zip," _
      & "       shippingLocState,shippingLocCountry,shippingZip " _
	  & "FROM   customer " _
	  & "WHERE  idCust = " & validSQL(idCust,"I")
set rsTemp = openRSexecute(mySQL)
if not rsTemp.eof then
	locCountry			= trim(rsTemp("locCountry")&"")
	locState			= trim(rsTemp("locState")&"")
	zip					= trim(rsTemp("zip")&"")
	shippingLocCountry	= trim(rsTemp("shippingLocCountry")&"")
	shippingLocState	= trim(rsTemp("shippingLocState")&"")
	shippingZip			= trim(rsTemp("shippingZip")&"")
	if len(shippingLocCountry & shippingLocState) = 0 then
		shippingLocCountry	= locCountry
		shippingLocState	= locState
		shippingZip			= zip
	end if
else
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrShipLoc)
end if
call closeRS(rsTemp)

'Get Shipping Zone
if len(shippingLocState) = 0 then
	mySQL = "SELECT locShipZone " _
		  & "FROM   locations " _
		  & "WHERE  locCountry = '" & validSQL(shippingLocCountry,"A") & "' " _
		  & "AND   (locState = '' OR locState IS NULL) "
else
	mySQL = "SELECT locShipZone " _
		  & "FROM   locations " _
		  & "WHERE  locCountry = '" & validSQL(shippingLocCountry,"A") & "' " _
		  & "AND    locState = '"   & validSQL(shippingLocState,"A")   & "' "
end if
set rsTemp = openRSexecute(mySQL)
if not rsTemp.eof then
	locShipZone = trim(rsTemp("locShipZone")&"")
else
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrShipZone)
end if
call closeRS(rsTemp)

'Get Total Price, Weight and Quantity of items with shipping charges
totalShipPrice	= 0
totalShipWeight	= 0
totalShipItems  = 0
mySQL="SELECT a.idCartRow, a.quantity " _
	& "FROM   cartRows a, products b " _
	& "WHERE  a.idOrder = " & validSQL(idOrder,"I") & " " _
	& "AND    b.idProduct = a.idProduct " _
	& "AND   (b.noShipCharge IS NULL " _
	& "OR     b.noShipCharge <> 'Y') "
set rsTemp = openRSexecute(mySQL)
do while not rsTemp.EOF
	totalShipPrice  = Cdbl(totalShipPrice  + cartTotal(idOrder,rsTemp("idCartRow")))
	totalShipWeight = Cdbl(totalShipWeight + cartRowWeight(idOrder,rsTemp("idCartRow")))
	totalShipItems  = Cdbl(totalShipItems  + rsTemp("quantity"))
	rsTemp.MoveNext
loop
call closeRS(rsTemp)

'For debug purposes, write shipping data as hidden HTML
Response.Write vbCrlf _
	& "<!-- " & vbCrlf _
	& "Ship Rate Price : " & totalShipPrice  & vbCrlf _
	& "Ship Rate Weight: " & totalShipWeight & vbCrlf _
	& "Ship Rate Items : " & totalShipItems  & vbCrlf _
	& "Ship Rate Zone  : " & locShipZone     & vbCrlf _
	& "-->" & vbCrlf

'-----------------
' Get Store Rates
'-----------------

'Check if order requires shipping charges
if totalShipItems = 0 then
	shipArray(0,0) = 0.00
	shipArray(0,1) = langGenNoShipCharge
else
	call calculateShipping(totalShipPrice,totalShipWeight,locShipZone)
end if

'------------------
' Get Online Rates
'------------------

'Online shipping parameters
dim UPSactive,UPSAccessID,UPSUserID,UPSPassword,UPSfromZip,UPSfromCntry,UPSpickupType,UPSpackType,UPSshipCode,UPSweightUnit,UPSallRates
dim USPSactive,USPSUserID,USPSPassword,USPSfromZip,USPSservice,USPSintNtl,USPSsize,USPSmachinable
dim CPactive,CPmerchantID,CPfromZip,CPsizeL,CPsizeW,CPsizeH

'Get online shipping parameters from database
mySQL = "SELECT configVar, configVal " _
	  & "FROM   storeAdmin " _
	  & "WHERE  adminType = 'S'"
set rsTemp = openRSexecute(mySQL)
do while not rsTemp.EOF
	select case trim(lCase(rsTemp("configVar")))
	
	'UPS
	case lCase("UPSactive")
		UPSactive			= rsTemp("configVal")
	case lCase("UPSAccessID")
		UPSAccessID			= rsTemp("configVal")
	case lCase("UPSUserID")
		UPSUserID			= rsTemp("configVal")
	case lCase("UPSPassword")
		UPSPassword			= rsTemp("configVal")
	case lCase("UPSfromZip")
		UPSfromZip			= rsTemp("configVal")
	case lCase("UPSfromCntry")
		UPSfromCntry		= rsTemp("configVal")
	case lCase("UPSpickupType")
		UPSpickupType		= rsTemp("configVal")
	case lCase("UPSpackType")
		UPSpackType			= rsTemp("configVal")
	case lCase("UPSshipCode")
		UPSshipCode			= rsTemp("configVal")
	case lCase("UPSweightUnit")
		UPSweightUnit		= rsTemp("configVal")
	case lCase("UPSallRates")
		UPSallRates			= rsTemp("configVal")
		
	'USPS
	case lCase("USPSactive")
		USPSactive			= rsTemp("configVal")
	case lCase("USPSUserID")
		USPSUserID			= rsTemp("configVal")
	case lCase("USPSPassword")
		USPSPassword		= rsTemp("configVal")
	case lCase("USPSfromZip")
		USPSfromZip			= rsTemp("configVal")
	case lCase("USPSservice")
		USPSservice			= rsTemp("configVal")
	case lCase("USPSintNtl")
		USPSintNtl			= rsTemp("configVal")
	case lCase("USPSsize")
		USPSsize			= rsTemp("configVal")
	case lCase("USPSmachinable")
		USPSmachinable		= rsTemp("configVal")
		
	'Canada Post
	case lCase("CPactive")
		CPactive			= rsTemp("configVal")
	case lCase("CPmerchantID")
		CPmerchantID		= rsTemp("configVal")
	case lCase("CPfromZip")
		CPfromZip			= rsTemp("configVal")
	case lCase("CPsizeL")
		CPsizeL				= rsTemp("configVal")
	case lCase("CPsizeW")
		CPsizeW				= rsTemp("configVal")
	case lCase("CPsizeH")
		CPsizeH				= rsTemp("configVal")
	end select
	rsTemp.MoveNext
loop
call closeRS(rsTemp)

'Store shipArray in the session object so that it can be accessed by
'the online rate routines which are invoked via "server.execute".
session(storeID & "shipArray") = shipArray

'Get UPS Online Rates
if UPSactive = "Y" then
	shipParms = array(UPSAccessID,UPSUserID,UPSPassword,UPSfromZip,UPSfromCntry,UPSpickupType,UPSpackType,UPSshipCode,UPSweightUnit,UPSallRates,totalShipWeight,shippingLocCountry,shippingZip)
	session(storeID & "shipParms") = shipParms
	server.Execute "_INCshipUPS_.asp"
end if

'Get USPS Online Rates
if USPSactive = "Y" then
	if UCase(shippingLocCountry) = "US" then	'US shipping
		shipParms = array(USPSUserID,USPSPassword,USPSfromZip,USPSservice,USPSsize,USPSmachinable,totalShipWeight,shippingLocCountry,shippingZip)
		session(storeID & "shipParms") = shipParms
		server.Execute "_INCshipUSPS_.asp"
	else
		if USPSintNtl = "Y" then				'International shipping
			shippingLocCountryDesc = getCountryDesc(shippingLocCountry)
			shipParms = array(USPSUserID,USPSPassword,totalShipWeight,shippingLocCountryDesc)
			session(storeID & "shipParms") = shipParms
			server.Execute "_INCshipUSPSi_.asp"
		end if
	end if
end if

'Get Canada Post Online Rates
if CPactive = "Y" then
	shipParms = array(CPmerchantID,CPfromZip,CPsizeL,CPsizeW,CPsizeH,totalShipWeight,shippingLocCountry,shippingLocState,shippingZip)
	session(storeID & "shipParms") = shipParms
	server.Execute "_INCshipCP_.asp"
end if

'Move session shipArray back to local shipArray
shipArray = session(storeID & "shipArray")

'Clean up session object
session(storeID & "shipArray")	= null
session(storeID & "shipParms")	= null

'------------------
' Get Custom Rates
'------------------
%>
<!--#include file="../UserMods/_INCship_.asp"-->
<%

'-------------------
' Check Rates Array
'-------------------

if isNull(shipArray(0,0)) or isEmpty(shipArray(0,0)) _
or isNull(shipArray(0,1)) or isEmpty(shipArray(0,1)) _
or len(shipArray(0,1)) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrNoShipRate)
end if
for f = 0 to UBound(shipArray)
	if  len(trim(shipArray(f,0))) > 0 _
	and len(trim(shipArray(f,1))) > 0  then
		if len(trim(shipArray(f,1))) > 100 _
		or not(isNumeric(shipArray(f,0))) then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvShipRate & " (" & f & ")")
		end if
	end if
next

'*********************************************************************
'*************   E N D - S H I P P I N G   R A T E S   ***************
'*********************************************************************

'Check if Payment is required for this Order
if (cartTotal(idOrder,0) + shipArray(0,0)) > 0 then
	paymentRequired = "Y"
else
	paymentRequired = "N"
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"-->
<%
'If there were errors, show message
if len(trim(arrayErrors)) > 0 then
	arrayErrors = split(LCase(arrayErrors),"|")
	Response.Write "<font color=red><i>" & langErrInvForm & "</i></font><br><br>"
else
	arrayErrors = array("")
end if
%>

<!-- Outer Table Cell -->
<table border="0" cellpadding="0" cellspacing="0" width=350><tr><td>

<!-- Inner Table -->
<table border="0" cellspacing="0" cellpadding="2" width="100%">

	<form METHOD="POST" name="shippingform" action="30_Ship_CC.asp">
	<input type=hidden name=formID          value="02">
	<input type=hidden name=paymentRequired value="<%=paymentRequired%>">
	<input type=hidden name=paymentType     value="<%=paymentType%>">
	
	<!-- Shipping Rate -->
	
	<tr>
		<td colspan=2 valign=middle class="CPpageHead">
			<table border=0 cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td nowrap align=left>
					<b><%=langGenMetShip%></b> <%=checkFieldError("shipDetails",arrayErrors)%>
				</td>
				<td nowrap align=right>
					<b><font color=#800000>[ <%=langGenStep%> 2 / 4 ]</font></b>
				</td>
			</tr>
			</table>
		</td>
	</tr>
    <TR> 
		<TD colspan=2><img src="../UserMods/misc_cleardot.gif" height=4 width=1><br></TD>
    </TR>
	<TR>
		<td nowrap colspan=2>
<%			
			'Handling Fee
			if totalShipWeight > 0 then
				handlingFeeTotal = handlingFeeAmt
			else
				handlingFeeTotal = 0.00
			end if
%>			<input type="hidden" name="handlingFeeTotal" value="<%=handlingFeeTotal%>">
			<!-- <%=langGenHandlingFeeMsg%> : <b><%=pCurrencySign & moneyS(handlingFeeTotal)%></b> -->
			<%=checkFieldError("handlingFeeTotal",arrayErrors)%><br><br>
<%
			'If only one shipping rate is available for selection, 
			'make sure that it's selected by default
			if len(trim(shipArray(1,0)&"")) = 0 then
				shipDetails = shipArray(0,0) & "|" & shipArray(0,1)
			end if

			'Display shipping rates
			if shipDisplayType = "0" then
				Response.Write "<select name=shipDetails size=1>"
			end if
			for f = 0 to Ubound(shipArray) 
				if len(trim(shipArray(f,0))) > 0 then
					if shipDisplayType = "0" then
%>
						<option <%=checkMatch(shipDetails,shipArray(f,0) & "|" & shipArray(f,1))%> value="<%=shipArray(f,0) & "|" & shipArray(f,1)%>"><%=shipArray(f,1) & " " & pCurrencySign & moneyS(shipArray(f,0))%></option>
<%
					else
%>
						<input type="radio" name="shipDetails" <%=replace(checkMatch(shipDetails,shipArray(f,0) & "|" & shipArray(f,1)),"selected","checked")%> value="<%=shipArray(f,0) & "|" & shipArray(f,1)%>"><%=shipArray(f,1) & " " & pCurrencySign & moneyS(shipArray(f,0))%><br>
<%
					end if
				end if
			next
			if shipDisplayType = "0" then
				Response.Write "</select>"
			end if
%>
		</td>
	</TR>
	
	<!-- Offline Credit Cards -->
<%
	'Get offline Credit Card info
	if lCase(paymentType) = "creditcard" and paymentRequired = "Y" then
%>
		<TR> 
			<TD COLSPAN="2">&nbsp;</TD>
		</TR>
		<TR>
			<td colspan=2 valign=middle class="CPpageHead">
				<b><%=langGenPayDetail%></b>
			</td>
		</TR>
		<TR> 
			<TD colspan=2><img src="../UserMods/misc_cleardot.gif" height=4 width=1><br></TD>
		</TR>
<%
		if demoMode = "Y" and len(cardNumber) = 0 then
			cardType     = "Visa"
			cardNumber   = "4111 1111 1111 1111"
			cardExpMonth = "6"
			cardExpYear  = "2005"
			cardVerify   = "123"
			cardName     = "Demo Card Name"
		end if
%>
		<TR> 
			<td nowrap><%=langGenCCtype & " " & checkFieldError("cardType",arrayErrors)%></td>
			<td>
				<select name="cardType" size=1>
					<option value="">
<%
					pCCTypeArr = split(pCCType,",")
					for f = 0 to Ubound(pCCTypeArr) 
						if len(trim(pCCTypeArr(f))) > 0 then
%>
						<option <%=checkMatch(cardType,pCCTypeArr(f))%> value="<%=pCCTypeArr(f)%>"><%=pCCTypeArr(f)%></option>
<%
						end if
					next
%>
				</select>
			</td>
		</TR>
		<TR> 
			<TD nowrap><%=langGenCCnumber & " " & checkFieldError("cardNumber",arrayErrors)%></TD>
			<TD>
				<input type=text name=cardNumber size=20 maxlength="20" value="<%=cardnumber%>">
			</TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenCCexpire & " " & checkFieldError("CardExpMonth",arrayErrors) & checkFieldError("CardExpYear",arrayErrors)%></TD>
			<TD>
				<select name="CardExpMonth">
					<option value=""><%=langGenMonth%></option>
<%
					for I = 1 to 12
%>
					<option <%=checkMatch(CardExpMonth,CStr(I))%> value="<%=I%>"><%=I%></option>
<%
					next
%>
				</select>
				&nbsp;/&nbsp;
				<select name="CardExpYear">
					<option value=""><%=langGenYear%></option>
<%
					for I = year(now()) to year(now()) + 10
%>
					<option <%=checkMatch(CardExpYear,CStr(I))%> value="<%=I%>"><%=I%></option>
<%
					next
%>
				</select>
			</TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenCCcvv & " " & checkFieldError("cardVerify",arrayErrors)%></TD>
			<TD>
				<input type=text name=cardVerify size=3 maxlength="3" value="<%=cardVerify%>">
			</TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenCCname & " " & checkFieldError("cardName",arrayErrors)%>
			</TD>
			<TD>
				<input type=text name=cardName size=25 maxlength="100" value="<%=cardName%>">
			</TD>
		</TR>
<%
	end if
%>
	<!-- Comments -->

	<TR> 
		<TD COLSPAN="2">&nbsp;</TD>
	</TR>
	<TR>
		<td colspan=2 valign=middle class="CPpageHead">
			<b><%=langGenAddComment%></b> <%=langGenCommentsHelp & " " & checkFieldError("generalComments",arrayErrors)%>
		</td>
	</TR>
	<TR> 
		<TD colspan=2><img src="../UserMods/misc_cleardot.gif" height=4 width=1><br></TD>
	</TR>
    <TR>
		<TD colspan=2>
			<textarea name="generalComments" rows=3 cols="35" wrap="soft"><%=generalComments%></textarea>
		</TD>
    </TR>
	
	
	<TR> 
			<TD nowrap><%=langGendeliverydate & " " & checkFieldError("deliverydate",arrayErrors)%></TD>
			<TD>
				<input type=text name=deliverydate size=20 maxlength="20" value="<%=deliverydate%>">
			</TD>
		</TR>
    
	<!-- Button -->

	<TR> 
		<TD colspan=2>&nbsp;</TD>
	</TR>
    <TR> 
		<TD colspan=2><input type="submit" name="Submit" value="<%=langGenNextBut%>"></TD>
    </TR>
	<TR>
		<TD colspan=2>&nbsp;</TD>
	</TR>

	</form>
	
</table>

<!-- End Outer Table Cell -->
</td></tr></table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%

call closedb()

'******************************************************************
'Calculate Store Rates from the database
'******************************************************************
sub calculateShipping(cartSubTotal, cartWeight, locShipZone)

	f = 0 'Initialize Counter

	'Get shipping rate records
	mySQL="SELECT a.addAmt,a.addPerc,b.shipDesc " _
	    & "FROM   shipRates a, shipMethod b " _
	    & "WHERE  a.idShipMethod = b.idShipMethod " _ 
	    & "AND    b.status = 'A' " _
	    & "AND    locShipZone = " & validSQL(locShipZone,"I") & " " _
	    & "AND   (addAmt IS NOT NULL OR addPerc IS NOT NULL) " _
	    & "AND  ((unitType='P' AND unitsFrom <= " & validSQL(cartSubTotal,"D") & " AND unitsTo >= " & validSQL(cartSubTotal,"D") & ") " _
	    & "OR    (unitType='W' AND unitsFrom <= " & validSQL(cartWeight,"D")   & " AND unitsTo >= " & validSQL(cartWeight,"D")   & ")) " _
	    & "ORDER BY a.idShipMethod " 
	set rsTemp = openRSexecute(mySQL)
	do while not rsTemp.eof

		'Get values from recordset
		addAmt   = rsTemp("addAmt")
		addPerc  = rsTemp("addPerc")
		shipDesc = trim(rsTemp("shipDesc")&"")

		'Calculate shipping based on fixed amount or percentage of 
		'order total (whichever is the greater value).
		if IsNull(addAmt) then
			addAmt = Round(((cartSubTotal * addPerc) / 100),2)
		else
			if not IsNull(addPerc) then
				if addAmt < ((cartSubTotal * addPerc) / 100) then
					addAmt = Round(((cartSubTotal * addPerc) / 100),2)
				end if
			end if
		end if
			
		'Move values into array, whilst making sure that we move 
		'the largest amount for each shipping method group.
		if f = 0 then '1st position in the array
			shipArray(f,0) = addAmt
			shipArray(f,1) = shipDesc
			f = f + 1
		else
			if lCase(shipDesc) = lCase(shipArray(f-1,1)) then
				if  shipArray(f-1,0) < addAmt then
					shipArray(f-1,0) = addAmt
				end if
			else
				shipArray(f,0) = addAmt
				shipArray(f,1) = shipDesc
				f = f + 1
			end if
		end if
		
		'Read next record
		rsTemp.movenext
			
	loop
	call closeRS(rsTemp)
	
end sub
'*************************************************************************
'Calculate Cart Row Weight
'*************************************************************************
function cartRowWeight(idOrder,idCartRow)
	dim mySQL, rsTemp, quantity, unitWeight, optionWeight
	cartRowWeight = CDbl(0)
	mySQL = "SELECT quantity, unitWeight, " _
		  & "      (SELECT SUM(optionWeight) " _
		  & "       FROM   cartRowsOptions " _
		  & "       WHERE  cartRowsOptions.idCartRow = cartRows.idCartRow) " _
		  & "       AS     optionWeight " _
		  & "FROM   cartRows " _
		  & "WHERE  idOrder   = " & validSQL(idOrder,"I") & " " _
		  & "AND    idCartRow = " & validSQL(idCartRow,"I") & " "
	set rsTemp = openRSexecute(mySQL)
	do while not rsTemp.eof
		quantity      = CDbl(emptyString(rsTemp("quantity"),"0"))
		unitWeight    = CDbl(emptyString(rsTemp("unitWeight"),"0"))
		optionWeight  = CDbl(emptyString(rsTemp("optionWeight"),"0"))
		cartRowWeight = cartRowWeight + (quantity * (unitWeight + optionWeight))
		rsTemp.movenext
	loop
	call closeRS(rsTemp)
end function
'*************************************************************************
'Check Credit Card Number (Test Number - 4111111111111111)
'*************************************************************************
function isCreditCard(cardNo)

	dim lCard, lC, cStat, temp, tempChar, i, d
	
	cardNo = trim(cardNo)
	cardNo = replace(cardNo," ","")
	cardNo = replace(cardNo,"-","")
	
	if isNumeric(cardNo) then 
		isCreditCard	= false 
		lCard			= len(cardNo) 
		lC				= right(cardNo,1) 
		cStat			= 0 
		for i=(lCard-1) to 1 step -1 
		    tempChar = mid(cardNo,i,1) 
		    d		 = CLng(tempChar) 
		    if lcard mod 2 = 1 then 
		        temp = d*(1+((i+1) mod 2)) 
		    else 
		        temp = d*(1+(i mod 2)) 
		    end if 
		    if temp < 10 then 
		        cStat = cStat + temp  
		    else 
		        cStat = cStat + temp - 9 
		    end if 
		next
		cStat = (10-(cStat mod 10)) mod 10 
		if CLng(lC) = cStat then 
			isCreditCard = true
		end if
	end if
end function
%>
