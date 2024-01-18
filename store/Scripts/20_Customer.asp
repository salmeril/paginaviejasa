<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Get user's name, address, shipping address.
'          : Validate Customer Info
'          : Create/Modify Customer Account
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
dim f, cont
dim arrayErrors
dim action
dim formID
dim countryArr
dim stateArr

'Customer
dim status
dim Name
dim LastName
dim CustomerCompany
dim Phone
dim Email
dim Password
dim Address
dim City
dim Zip
dim locState
dim locState2
dim locCountry
dim paymentType
dim shippingName
dim shippingLastName
dim shippingPhone
dim shippingAddress
dim ShippingCity
dim shippingZip
dim shippingLocState
dim shippingLocState2
dim shippingLocCountry
dim futureMail
dim taxExempt

'Locations
dim locName

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

'Check action indicator
action = lCase(Request.QueryString("action"))
if len(action) = 0 then
	action = lCase(Request.Form("action"))
end if
if  action <> "newacc" _
and action <> "modify" _
and action <> "save" _
and action <> "checkout" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrAction)
end if

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

'If Modify, do some validations.
if action = "modify" then
	
	'Check that Customer is logged on
	if isNull(idCust) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrNotLoggedIn)
	end if
	
end if

'Get Form ID
formID = trim(Request.Form("formID"))

'Before we display the form for the first time, do some checks
if formID = "" then

	'Default Country Code
	locCountry = defaultCountryCode
	ShippinglocCountry = defaultCountryCode


	'Check if user is already logged on
	if not isNull(idCust) then
	
		'Retrieve field values from DB
		mySQL = "SELECT Name,LastName,CustomerCompany,Phone,Email," _
			  & "		Password,Address,City,Zip,locCountry,locState," _
			  & "		locState2,paymentType,shippingName,shippingLastName," _
			  & "       shippingPhone,shippingAddress,ShippingCity," _
			  & "       shippingZip,shippingLocCountry,shippingLocState," _
			  & "       shippingLocState2,futureMail " _
			  & "FROM   customer " _
			  & "WHERE  idCust = " & validSQL(idCust,"I")
		set rsTemp = openRSexecute(mySQL)
		if not rsTemp.EOF then
			Name				= trim(rstemp("name")&"")
			LastName			= trim(rstemp("LastName")&"")
			CustomerCompany		= trim(rstemp("CustomerCompany")&"")
			Phone				= trim(rstemp("Phone")&"")
			Email				= trim(rstemp("Email")&"")
			Password			= trim(EnDeCrypt(Hex2Ascii(rstemp("Password")),rc4Key)&"")
			Address				= trim(rstemp("Address")&"")
			City				= trim(rstemp("City")&"")
			Zip					= trim(rstemp("Zip")&"")
			locState			= trim(rstemp("locState")&"")
			locState2			= trim(rstemp("locState2")&"")
			locCountry			= trim(rstemp("locCountry")&"")
			paymentType			= trim(rstemp("paymentType")&"")
			shippingName		= trim(rstemp("shippingName")&"")
			shippingLastName	= trim(rstemp("shippingLastName")&"")
			shippingPhone		= trim(rstemp("shippingPhone")&"")
			shippingAddress		= trim(rstemp("shippingAddress")&"")
			ShippingCity		= trim(rstemp("ShippingCity")&"")
			shippingZip			= trim(rstemp("shippingZip")&"")
			shippingLocState	= trim(rstemp("shippingLocState")&"")
			shippingLocState2	= trim(rstemp("shippingLocState2")&"")
			shippingLocCountry	= trim(rstemp("shippingLocCountry")&"")
			futureMail			= trim(rstemp("futureMail")&"")
		else
			'No Customer Record on DB (which is highly unlikely because
			'Customer record has already been tested in sessionCust()
			'at the beginning of this script).
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvCustAcc)
		end if
		call closeRS(rsTemp)
		
	end if
	
end if
	
'Check if the Customer clicked the "Next" button
if formID = "01" then
	Name				= validHTML(request.form("name"))
	LastName			= validHTML(request.form("LastName"))
	CustomerCompany		= validHTML(request.form("CustomerCompany"))
	Phone				= validHTML(request.form("Phone"))
	Email				= validHTML(request.form("Email"))
	Password			= validHTML(request.form("Password"))
	Address				= validHTML(request.form("Address"))
	City				= validHTML(request.form("City"))
	Zip					= validHTML(request.form("Zip"))
	locState			= validHTML(request.form("locState"))
	locState2			= validHTML(request.form("locState2"))
	locCountry			= validHTML(request.form("locCountry"))
	paymentType			= validHTML(request.form("paymentType"))
	shippingName		= validHTML(request.form("shippingName"))
	shippingLastName	= validHTML(request.form("shippingLastName"))
	shippingPhone		= validHTML(request.form("shippingPhone"))
	shippingAddress		= validHTML(request.form("shippingAddress"))
	ShippingCity		= validHTML(request.form("ShippingCity"))
	shippingZip			= validHTML(request.form("shippingZip"))
	shippingLocState	= validHTML(request.form("shippingLocState"))
	shippingLocState2	= validHTML(request.form("shippingLocState2"))
	shippingLocCountry	= validHTML(request.form("shippingLocCountry"))
	futureMail			= validHTML(request.form("futureMail"))
	
	'Name
	if len(name) = 0 then
		arrayErrors = arrayErrors & "|name"
	end if
	
	'LastName
	if len(lastname) = 0 then
		arrayErrors = arrayErrors & "|lastname"
	end if
	
	'Phone
	if len(phone) = 0 then
		arrayErrors = arrayErrors & "|phone"
	else
		if invalidChar(phone,2,"- +().") then
			arrayErrors = arrayErrors & "|phone"
		end if
	end if
	
	'Email
	if len(email) = 0 then
		arrayErrors = arrayErrors & "|email"
	else
		if inStr(email,"@") = 0 or inStr(email,".") = 0 then
			arrayErrors = arrayErrors & "|email"
		end if
		if invalidChar(Email,1,"@.-_") then
			arrayErrors = arrayErrors & "|email"
		end if
	end if
	
	'Password
	if len(password) = 0 then
		arrayErrors = arrayErrors & "|password"
	else
		if invalidChar(Password,1,"") then
			arrayErrors = arrayErrors & "|password"
		end if
	end if
		
	paymentType = payDefault
	'PaymentType
	if len(paymentType) = 0 then
		arrayErrors = arrayErrors & "|paymenttype"
	end if

	
	'Future Mail Indicator
	if futureMail <> "Y" then
		futureMail = "N"
	end if
	
	'Address
	if len(address) = 0 then
		arrayErrors = arrayErrors & "|address"
	end if
	
	'City
	if len(city) = 0 then
		arrayErrors = arrayErrors & "|city"
	end if
	
	'Zip
	if len(zip) = 0 then
		arrayErrors = arrayErrors & "|zip"
	end if
	
	'State/Prov/Country
	if len(locCountry) = 0 then
		arrayErrors = arrayErrors & "|locState"
		arrayErrors = arrayErrors & "|locCountry"
	else
		if not validLoc(locState,locCountry) then
			arrayErrors = arrayErrors & "|locState"
			arrayErrors = arrayErrors & "|locCountry"
		end if
	end if
	
	'State/Province 2
	if len(locState) > 0 and len(locState2) > 0 then
		arrayErrors = arrayErrors & "|locState2"
	end if
	
	'Shipping
	if len(shippingName & shippingLastName & shippingPhone & shippingAddress & shippingCity & shippingZip & shippingLocCountry) > 0 then
		'Ship Name
		if len(shippingName) = 0 then
			arrayErrors = arrayErrors & "|shippingName"
		end if
		'Ship Last Name
		if len(shippingLastName) = 0 then
			arrayErrors = arrayErrors & "|shippingLastName"
		end if
		'shippingPhone
		if len(shippingPhone) = 0 then
			arrayErrors = arrayErrors & "|shippingPhone"
		else
			if invalidChar(shippingPhone,2,"- +().") then
				arrayErrors = arrayErrors & "|shippingPhone"
			end if
		end if
		'Ship Address
		if len(shippingAddress) = 0 then
			arrayErrors = arrayErrors & "|shippingAddress"
		end if
		'Ship City
		if len(shippingCity) = 0 then
			arrayErrors = arrayErrors & "|shippingCity"
		end if
		'Ship Zip
		if len(shippingZip) = 0 then
			arrayErrors = arrayErrors & "|shippingZip"
		end if
		'Ship State/Prov/Country
		if len(shippingLocCountry) = 0 then
			arrayErrors = arrayErrors & "|shippingLocState"
			arrayErrors = arrayErrors & "|shippingLocCountry"
		else
			if not validLoc(shippingLocState,shippingLocCountry) then
				arrayErrors = arrayErrors & "|shippingLocState"
				arrayErrors = arrayErrors & "|shippingLocCountry"
			end if
		end if
		'Ship State/Province 2
		if len(shippingLocState) > 0 and len(shippingLocState2) > 0 then
			arrayErrors = arrayErrors & "|shippingLocState2"
		end if
	end if
	
	'There were no errors
	if len(trim(arrayErrors)) = 0 then
	
		'Check for duplicate email address
		mySQL = "SELECT idCust " _
			  & "FROM   customer " _
			  & "WHERE  email = '" & validSQL(email,"A") & "' "
		if not isNull(idCust) then
			mySQL = mySQL & "AND idCust <> " & validSQL(idCust,"I")
		end if
		set rsTemp = openRSexecute(mySQL)
		if not rsTemp.EOF then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrEmailUsed)
		end if
		call closeRS(rsTemp)
	
		'Create empty Customer Record
		if isNull(idCust) then
			set rsTemp = openRSopen("customer",adUseServer,adOpenKeySet,adLockOptimistic,adCmdTable,0)
			rsTemp.AddNew
			rsTemp("status")		= "A" 'Active Customer Record
			rsTemp("dateCreated")	= now()
			rsTemp("dateCreatedInt")= dateInt(now())
			rsTemp("taxExempt")     = "N" 'Default to No
			rsTemp.update 
			session(storeID & "idCust") = rsTemp("idCust")
			idCust						= rsTemp("idCust")
			call closeRS(rsTemp)
		end if

		'Update the customer record
		mySQL = "UPDATE customer SET " _
			&   "[name]				= '" & validSQL(name,"A") & "', " _
			&   "LastName			= '" & validSQL(LastName,"A") & "', " _
			&   "CustomerCompany	= '" & validSQL(CustomerCompany,"A") & "', " _
			&   "Phone				= '" & validSQL(Phone,"A") & "', " _
			&   "Email				= '" & validSQL(Email,"A") & "', " _
			&   "[Password]			= '" & validSQL(Ascii2Hex(EnDeCrypt(lCase(Password),rc4Key)),"A") & "', " _
			&   "Address			= '" & validSQL(Address,"A") & "', " _
			&   "City				= '" & validSQL(City,"A") & "', " _
			&   "Zip				= '" & validSQL(Zip,"A") & "', " _
			&   "locState			= '" & validSQL(locState,"A") & "', " _
			&   "locState2			= '" & validSQL(locState2,"A") & "', " _
			&   "locCountry			= '" & validSQL(locCountry,"A") & "', " _
			&   "paymentType		= '" & validSQL(paymentType,"A") & "', " _
			&   "shippingName		= '" & validSQL(shippingName,"A") & "', " _
			&   "shippingLastName	= '" & validSQL(shippingLastName,"A") & "', " _
			&   "shippingPhone		= '" & validSQL(shippingPhone,"A") & "', " _
			&   "ShippingAddress	= '" & validSQL(ShippingAddress,"A") & "', " _
			&   "ShippingCity		= '" & validSQL(ShippingCity,"A") & "', " _
			&   "shippingZip		= '" & validSQL(shippingZip,"A") & "', " _
			&   "shippingLocState	= '" & validSQL(shippingLocState,"A") & "', " _
			&   "shippingLocState2	= '" & validSQL(shippingLocState2,"A") & "', " _
			&   "shippingLocCountry	= '" & validSQL(shippingLocCountry,"A") & "', " _
			&   "futureMail			= '" & validSQL(futureMail,"A") & "' " _
			&   "WHERE idCust		=  " & validSQL(idCust,"I")
		set rsTemp = openRSexecute(mySQL)
		call closeRS(rsTemp)
		
		'Update cartHead with some info (if possible)
		if not(isNull(idOrder)) then
			mySQL = "UPDATE cartHead SET " _
				  & "idCust        = "  & validSQL(idCust,"I")	 & ", " _
				  & "[Name]        = '" & validSQL(Name,"A")	 & "'," _
				  & "LastName      = '" & validSQL(LastName,"A") & "'," _
				  & "Address	   = '" & validSQL(Address,"A")	 & "' " _
				  & "WHERE idOrder = "  & validSQL(idOrder,"I")	 & " "
			set rsTemp = openRSexecute(mySQL)
			call closeRS(rsTemp)
		end if
		
		'Forward to next page
		select case action
		case "newacc"
			Response.Redirect "custListOrders.asp"
		case "modify"
			Response.Redirect "custListOrders.asp"
		case "save"
			call saveCart(idOrder,idCust)
			Response.Redirect "custListOrders.asp"
		case "checkout"
			Response.Redirect "30_Ship_CC.asp"
		end select
	
	end if
	
end if

'Determine default Payment Type
if len(paymentType) = 0 then
	paymentType = payDefault
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"-->
<%
showForm()
%>
<!--#include file="../UserMods/_INCbottom_.asp"-->
<%

call closedb()

'**********************************************************************
'Display the Customer Info Form
'**********************************************************************
sub showForm()

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

	<form METHOD="POST" name="orderform" action="20_Customer.asp">
	<input type=hidden name=action value="<%=action%>">
	<input type=hidden name=formID value="01">
	
	<!-- General Info -->
	
	<tr>
		<td colspan=2 valign=middle class="CPpageHead">
			<table border=0 cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td nowrap align=left>
					<b><%=langGenCustInfo%></b>
				</td>
				<td nowrap align=right>
<%
					'Display an appropriate heading
					select case action
					case "newacc", "save"
						Response.Write "<b><font color=#800000>" & langGenNewAcc & "</font></b>"
					case "modify"
						Response.Write "<b><font color=#800000>" & langGenModAcc & "</font></b>"
					case else
						Response.Write "<b><font color=#800000>[ " & langGenStep & " 1 / 4 ]</font></b>"
					end select
%>
				</td>
			</tr>
			</table>
		</td>
	</tr>
    <TR> 
		<TD colspan=2><img src="../UserMods/misc_cleardot.gif" height=4 width=1><br></TD>
    </TR>
    <TR> 
		<TD nowrap><%=langGenName & " " & checkFieldError("name",arrayErrors)%></TD>
		<TD><input type=text name=name size=30 maxlength="70" value="<%=name%>"></TD>
    </TR>
     <TR> 
		<TD nowrap><%=langGenLastName & " " & checkFieldError("lastName",arrayErrors)%></TD>
		<TD><input type=text name=lastName size=30 maxlength="50" value="<%=lastname%>"></TD>
    </TR>   
     <TR> 
		<TD nowrap><%=langGenCompany & " " & checkFieldError("customerCompany",arrayErrors)%></TD>
		<TD><input type=text name=customerCompany size=30 maxlength="50" value="<%=customercompany%>"></TD>
    </TR>   
    <TR> 
		<TD nowrap><%=langGenPhone & " " & checkFieldError("phone",arrayErrors)%></TD>
		<TD><input type=text name=phone size=30 maxlength="30" value="<%=phone%>"></TD>
    </TR>
    <TR> 
		<TD nowrap><%=langGenEmail & " " & checkFieldError("email",arrayErrors)%></TD>
		<TD><input type=text name=email size=30 maxlength="50" value="<%=email%>"></TD>
    </TR>
    <TR> 
		<TD nowrap><%=langGenPassword & " " & checkFieldError("password",arrayErrors)%></TD>
		<TD><input type=password name=password size=10 maxlength="10" value="<%=password%>"></TD>
    </TR>
	<!--<TR> 
		<td nowrap><%=langGenPayment & " " & checkFieldError("paymentType",arrayErrors)%></td>
		<td>
			<select name="paymentType" size=1>
				<option value="">------- Select -------</option>
<%
				if pMailIn = -1 then
					Response.Write "<option value=""MailIn"" " & checkMatch(paymentType,"mailin") & ">" & paymentMsg("mailin",1,"")
				end if
				if payCallIn = -1 then
					Response.Write "<option value=""CallIn"" " & checkMatch(paymentType,"callin") & ">" & paymentMsg("callin",1,"")
				end if
				if payFaxIn = -1 then
					Response.Write "<option value=""FaxIn"" " & checkMatch(paymentType,"faxin") & ">" & paymentMsg("faxin",1,"")
				end if
				if payCOD = -1 then
					Response.Write "<option value=""COD"" " & checkMatch(paymentType,"cod") & ">" & paymentMsg("cod",1,"")
				end if
				if pCreditCard = -1 then
					Response.Write "<option value=""CreditCard"" " & checkMatch(paymentType,"creditcard") & ">" & paymentMsg("creditcard",1,"")
				end if
				if pPayPal = -1 then
					Response.Write "<option value=""PayPal"" " & checkMatch(paymentType,"paypal") & ">" & paymentMsg("paypal",1,"")
				end if
				if TwoCheckOut = -1 then
					Response.Write "<option value=""2CheckOut"" " & checkMatch(paymentType,"2checkout") & ">" & paymentMsg("2checkout",1,"")
				end if
				if pAuthNetFrontEnd = -1 then
					Response.Write "<option value=""AuthorizeNet"" " & checkMatch(paymentType,"authorizenet") & ">" & paymentMsg("authorizenet",1,"")
				end if
				if payCustom = -1 then
					Response.Write "<option value=""Custom"" " & checkMatch(paymentType,"custom") & ">" & paymentMsg("custom",1,"")
				end if
%>
			</select>
			&nbsp;<a href="<%=urlNonSSL%>termsAndCond.asp" onClick='window.open("<%=urlNonSSL%>termsAndCond.asp","generalConditions","width=300,height=300,resizable=1,scrollbars=1");return false;' target="_blank"><%=langGenLearnMore%></a><br>
		</td>
	</TR>-->
    <TR> 
		<TD colspan=2>
			<input type=checkbox name=futureMail value="Y" <%if futureMail="Y" then Response.Write " checked " end if%>> <%=langGenNotifyMsg%>
		</TD>
    </TR>
    
	<!-- Billing Address -->
	
    <TR> 
		<TD COLSPAN="2">&nbsp;</TD>
    </TR>
	<TR>
		<td colspan=2 valign=middle class="CPpageHead">
			<b><%=langGenBillAddr%></b>
		</td>
    </TR>
    <TR> 
		<TD colspan=2><img src="../UserMods/misc_cleardot.gif" height=4 width=1><br></TD>
    </TR>
    <TR> 
		<TD nowrap><%=langGenAddress & " " & checkFieldError("address",arrayErrors)%></TD>
		<TD><input type=text name=address size=30 maxlength="70" value="<%=address%>"></TD>
    </TR>
    <TR> 
		<TD nowrap><%=langGenCity & " " & checkFieldError("city",arrayErrors)%></TD>
		<TD><input type=text name=city size=30 maxlength="50" value="<%=city%>"></TD>
    </TR>
    <TR> 
		<TD nowrap><%=langGenZip & " " & checkFieldError("zip",arrayErrors)%></TD>
		<TD><input type=text name=zip size=10 maxlength="10" value="<%=zip%>"></TD>
    </TR>
    <TR> 
		<TD nowrap><%=langGenState & " " & checkFieldError("locState",arrayErrors)%></TD>
		<TD><%listStates "locState",locState%></TD>
    </TR>
    <TR> 
		<TD>&nbsp;</TD>
		<TD><%=langGenStateAlt%></TD>
    </TR>
    <TR> 
		<TD align=right><%=checkFieldError("locState2",arrayErrors)%>&nbsp;&nbsp;</TD>
		<TD><input type=text name=locState2 size=30 maxlength="100" value="<%=locState2%>"></TD>
    </TR>
    <TR> 
		<TD nowrap><%=langGenCountry & " " & checkFieldError("locCountry",arrayErrors)%></TD>
		<TD><%listCountries "locCountry",locCountry%></TD>
    </TR>
    
	<!-- Shipping Address -->
<%
	'Check if we must show shipping address fields
	if allowShipAddr = -1 then
%>
		<TR> 
			<TD COLSPAN="2">&nbsp;</TD>
		</TR>
		<TR>
			<td colspan=2 valign=middle class="CPpageHead">
				<b><%=langGenShipAddr%></b> *
			</td>
		</TR>
		<TR> 
			<TD COLSPAN="2">
				<i>* <%=langGenShipAddrOpt%></i>
			</TD>
		</TR>
	    <TR> 
			<TD colspan=2><img src="../UserMods/misc_cleardot.gif" height=4 width=1><br></TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenName & " " & checkFieldError("shippingName",arrayErrors)%></TD>
			<TD><input type=text name=shippingName size=30 maxlength="70" value="<%=shippingName%>"></TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenLastName & " " & checkFieldError("shippingLastName",arrayErrors)%></TD>
			<TD><input type=text name=shippingLastName size=30 maxlength="70" value="<%=shippingLastName%>"></TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenPhone & " " & checkFieldError("shippingPhone",arrayErrors)%></TD>
			<TD><input type=text name=shippingPhone size=30 maxlength="30" value="<%=shippingPhone%>"></TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenAddress & " " & checkFieldError("shippingAddress",arrayErrors)%></TD>
			<TD><input type=text name=shippingAddress size=30 maxlength="70" value="<%=shippingaddress%>"></TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenCity & " " & checkFieldError("shippingCity",arrayErrors)%></TD>
			<TD><input type=text name=shippingCity size=30 maxlength="50" value="<%=shippingcity%>"></TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenZip & " " & checkFieldError("shippingZip",arrayErrors)%></TD>
			<TD><input type=text name=shippingZip size=10 maxlength="10" value="<%=shippingzip%>"></TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenState & " " & checkFieldError("shippinglocState",arrayErrors)%></TD>
			<TD><%listStates "shippinglocState",shippinglocState%></TD>
		</TR>
		<TR> 
			<TD>&nbsp;</TD>
			<TD><%=langGenStateAlt%></TD>
		</TR>
		<TR> 
			<TD align=right><%=checkFieldError("shippinglocState2",arrayErrors)%>&nbsp;&nbsp;</TD>
			<TD><input type=text name=shippinglocState2 size=30 maxlength="100" value="<%=shippinglocState2%>"></TD>
		</TR>
		<TR> 
			<TD nowrap><%=langGenCountry & " " & checkFieldError("shippinglocCountry",arrayErrors)%></TD>
			<TD><%listCountries "shippinglocCountry",shippinglocCountry%></TD>
		</TR>
<%
	'If shipping address must NOT be shown, substitute visible form
	'variables with hidden form variables.
	else
%>
		<input type=hidden name=shippingName value="">
		<input type=hidden name=shippingLastName value="">
		<input type=hidden name=shippingPhone value="">
		<input type=hidden name=shippingAddress value="">
		<input type=hidden name=shippingCity value="">
		<input type=hidden name=shippingZip value="">
		<input type=hidden name=shippinglocState value="">
		<input type=hidden name=shippinglocState2 value="">
		<input type=hidden name=shippinglocCountry value="">
<%
	end if
%>
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
	
</TABLE>

<!-- End Outer Table Cell -->
</td></tr></table>

<br>

<%
end sub
'*********************************************************************
'Validate State & Country Code combination
'*********************************************************************
function validLoc(locState,locCountry)
	
	dim mySQL, rsTemp
	
	'Assume error until proven otherwise
	validLoc = false
	
	'Do basic checks
	if len(locCountry) = 0 then
		exit function		
	end if

	'Do database checks
	if len(locState) = 0 then
		'Check if State/Province is required for the Country
		mySQL = "SELECT COUNT(*) as recCount " _
			  & "FROM   locations " _
			  & "WHERE  locCountry = '" & validSQL(locCountry,"A") & "' "
		set rsTemp = openRSexecute(mySQL)
		if rsTemp("recCount") <= 1 then	'No states defined for Country
			validLoc = true
		end if
		call closeRS(rsTemp)
	else
		'Check State/Country Combo is valid
		mySQL = "SELECT locState " _
			  & "FROM   locations " _
			  & "WHERE  locState = '"   & validSQL(locState,"A")   & "' " _
			  & "AND    locCountry = '" & validSQL(locCountry,"A") & "' "
		set rsTemp = openRSexecute(mySQL)
		if not rsTemp.EOF then			'Valid State/Country combo
			validLoc = true
		end if
		call closeRS(rsTemp)
	end if

end function
'*********************************************************************
'Create drop-down for states
'*********************************************************************
sub listStates(fieldName,fieldVal)

	dim mySQL,rsTemp,row
	
	'If state array doesn't already exist, create it from DB.
	if not isArray(stateArr) then
		mySQL="SELECT a.locName,a.locCountry,a.locState " _
		    & "FROM   locations a " _
			& "WHERE  a.locStatus = 'A' " _
			& "AND    NOT(a.locState IS NULL OR a.locState='') " _
			& "AND    EXISTS(SELECT b.idLocation " _
			& "              FROM   locations b " _
			& "              WHERE  b.locCountry = a.locCountry " _
			& "              AND    b.locStatus  = 'A' " _
			& "              AND   (b.locState IS NULL OR b.locState='')) " _
		    & "ORDER BY a.locCountry,a.locName"
		set rsTemp = openRSexecute(mySQL)
		if not rsTemp.EOF then
			stateArr = rsTemp.getRows()
		end if
		call closeRS(rsTemp)
	end if
%>
	<SELECT name="<%=fieldName%>" size=1>
<%
		if isArray(stateArr) then
%>
			<OPTION value=""></OPTION>
<%
			for row = 0 to UBound(stateArr,2)
%>
			<OPTION VALUE="<%=stateArr(2,row)%>" <%=checkMatch(fieldVal,stateArr(2,row))%>><%=trim(stateArr(0,row)) & " (" & stateArr(1,row) & ")"%></OPTION>
<%
			next
		else
%>
			<OPTION value=""><%=langGenNotApplicable%></OPTION>
<%
		end if
%>
	</SELECT>
<%
end sub
'*********************************************************************
'Create drop-down for countries
'*********************************************************************
sub listCountries(fieldName,fieldVal)

	dim mySQL,rsTemp,row
	
	'If country array doesn't already exist, create it from DB.
	if not isArray(countryArr) then
		mySQL="SELECT locName,locCountry " _
		    & "FROM   locations " _
		    & "WHERE  locStatus = 'A' " _
			& "AND    NOT(locCountry IS NULL OR locCountry='') " _
			& "AND    (locState IS NULL OR locState='') " _
		    & "ORDER BY locName"
		set rsTemp = openRSexecute(mySQL)
		countryArr = rsTemp.getRows()
		call closeRS(rsTemp)
	end if
%>
	<SELECT name="<%=fieldName%>" size=1>
<%
		if isArray(countryArr) then
%>
			<OPTION value=""></OPTION>
<%
			for row = 0 to UBound(countryArr,2)
%>
			<OPTION VALUE="<%=countryArr(1,row)%>" <%=checkMatch(fieldVal,countryArr(1,row))%>><%=trim(countryArr(0,row))%></OPTION>
<%
			next
		else
%>
			<OPTION value=""><%=langGenNotApplicable%></OPTION>
<%
		end if
%>
	</SELECT>
<%
end sub
%>
