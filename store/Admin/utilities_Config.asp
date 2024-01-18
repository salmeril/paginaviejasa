<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Modify Store General Settings
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
const adminLevel = 0
%>
<!--#include file="../Config/config.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="../Scripts/_INCappFunctions_.asp"-->
<%
'Config Variables
dim urlNonSSL,urlSSL,pDownloadDir,pImagesDir,mailComp
dim pSMTPServer,pEmailSales,pEmailAdmin,pCompany,pCatalogOnly
dim pMaxCartQty,prodViewLayout,pHidePricingZero,pMaxItemQty
dim pMinCartAmount,pMaxItemsPerPage,pOrderPrefix,pCurrencySign
dim pStoreLCID,pShowStockView,pMailIn,pPayPal,payPalMemberID
dim pCreditCard,pCCType,pAuthNet,authNetLogin,authNetCurrCode
dim payMsgMailIn,payMsgCreditCard,payMsgPayPal,payMsgOther
dim payMsgNotReq,shipDisplayType,TwoCheckOut,TwoCheckOutSID
dim payMsgTwoCheckOut,pEmailFriendSec,pMaxDownloadHours
dim pMaxDownloadCount,payDefault,pAuthNetFrontEnd,payMsgAuthNet
dim pCompanyAddr,TwoCheckoutMD5,pHideAddStockLevel,payCustom
dim payMsgCustom,taxOnShipping,allowShipAddr,defaultCountryCode
dim listViewLayout,payCallIn,payFaxIn,PayCOD,statUpdPending
dim payMsgCallIn,payMsgFaxIn,payMsgCOD,taxBillOrShip
dim handlingFeeAmt,handlingFeeTax,payPalCurrCode,authNetTxKey

'Database variables
dim mySQL, cn, rs
'*************************************************************************

'Open Database
call openDb()

'Get current configuration settings from database
mySQL = "SELECT configVar, configVal " _
	  & "FROM   storeAdmin " _
	  & "WHERE  adminType = 'C'"
set rs = openRSexecute(mySQL)
do while not rs.EOF

	select case trim(lCase(rs("configVar")))
	case lCase("urlNonSSL")
		urlNonSSL			= rs("configVal")
	case lCase("urlSSL")
		urlSSL				= rs("configVal")
	case lCase("pDownloadDir")
		pDownloadDir		= rs("configVal")
	case lCase("pImagesDir")
		pImagesDir			= rs("configVal")
	case lCase("mailComp")
		mailComp			= rs("configVal")
	case lCase("pSMTPServer")
		pSMTPServer			= rs("configVal")
	case lCase("pEmailSales")
		pEmailSales			= rs("configVal")
	case lCase("pEmailAdmin")
		pEmailAdmin			= rs("configVal")
	case lCase("pCompany")
		pCompany			= rs("configVal")
	case lCase("pCatalogOnly")
		pCatalogOnly		= rs("configVal")
	case lCase("pHidePricingZero")
		pHidePricingZero	= rs("configVal")
	case lCase("pMaxCartQty")
		pMaxCartQty			= rs("configVal")
	case lCase("pMaxItemQty")
		pMaxItemQty			= rs("configVal")
	case lCase("pMinCartAmount")
		pMinCartAmount		= rs("configVal")
	case lCase("pMaxItemsPerPage")
		pMaxItemsPerPage	= rs("configVal")
	case lCase("pOrderPrefix")
		pOrderPrefix		= rs("configVal")
	case lCase("pCurrencySign")
		pCurrencySign		= rs("configVal")
	case lCase("pStoreLCID")
		pStoreLCID			= rs("configVal")
	case lCase("pShowStockView")
		pShowStockView		= rs("configVal")
	case lCase("pMailIn")
		pMailIn				= rs("configVal")
	case lCase("pPayPal")
		pPayPal				= rs("configVal")
	case lCase("payPalMemberID")
		payPalMemberID		= rs("configVal")
	case lCase("TwoCheckOut")
		TwoCheckOut			= rs("configVal")
	case lCase("TwoCheckOutSID")
		TwoCheckOutSID		= rs("configVal")
	case lCase("pCreditCard")
		pCreditCard			= rs("configVal")
	case lCase("pCCType")
		pCCType				= rs("configVal")
	case lCase("pAuthNet")
		pAuthNet			= rs("configVal")
	case lCase("authNetLogin")
		authNetLogin		= rs("configVal")
	case lCase("authNetCurrCode")
		authNetCurrCode		= rs("configVal")
	case lCase("payMsgMailIn")
		payMsgMailIn		= rs("configVal")
	case lCase("payMsgCreditCard")
		payMsgCreditCard	= rs("configVal")
	case lCase("payMsgPayPal")
		payMsgPayPal		= rs("configVal")
	case lCase("payMsgTwoCheckOut")
		payMsgTwoCheckOut	= rs("configVal")
	case lCase("payMsgOther")
		payMsgOther			= rs("configVal")
	case lCase("payMsgNotReq")
		payMsgNotReq		= rs("configVal")
	case lCase("pEmailFriendSec")
		pEmailFriendSec		= rs("configVal")
	case lCase("pMaxDownloadHours")
		pMaxDownloadHours	= rs("configVal")
	case lCase("pMaxDownloadCount")
		pMaxDownloadCount	= rs("configVal")
	case lCase("payDefault")
		payDefault			= rs("configVal")
	case lCase("pAuthNetFrontEnd")
		pAuthNetFrontEnd	= rs("configVal")
	case lCase("pCompanyAddr")
		pCompanyAddr		= rs("configVal")
	case lCase("payMsgAuthNet")
		payMsgAuthNet		= rs("configVal")
	case lCase("TwoCheckoutMD5")
		TwoCheckoutMD5		= rs("configVal")
	case lCase("pHideAddStockLevel")
		pHideAddStockLevel	= rs("configVal")
	case lCase("payCustom")
		payCustom			= rs("configVal")
	case lCase("payMsgCustom")
		payMsgCustom		= rs("configVal")
	case lCase("taxOnShipping")
		taxOnShipping		= rs("configVal")
	case lCase("allowShipAddr")
		allowShipAddr		= rs("configVal")
	case lCase("prodViewLayout")
		prodViewLayout		= rs("configVal")
	case lCase("shipDisplayType")
		shipDisplayType		= rs("configVal")
	case lCase("defaultCountryCode")
		defaultCountryCode	= rs("configVal")
	case lCase("payCallIn")
		payCallIn			= rs("configVal")
	case lCase("payFaxIn")
		payFaxIn			= rs("configVal")
	case lCase("payCOD")
		payCOD				= rs("configVal")
	case lCase("payMsgCallIn")
		payMsgCallIn		= rs("configVal")
	case lCase("payMsgFaxIn")
		payMsgFaxIn			= rs("configVal")
	case lCase("payMsgCOD")
		payMsgCOD			= rs("configVal")
	case lCase("listViewLayout")
		listViewLayout		= rs("configVal")
	case lCase("taxBillOrShip")
		taxBillOrShip		= rs("configVal")
	case lCase("statUpdPending")
		statUpdPending		= rs("configVal")
	case lCase("handlingFeeAmt")
		handlingFeeAmt		= rs("configVal")
	case lCase("handlingFeeTax")
		handlingFeeTax		= rs("configVal")
	case lCase("payPalCurrCode")
		payPalCurrCode		= rs("configVal")
	case lCase("authNetTxKey")
		authNetTxKey		= rs("configVal")
	end select

	rs.MoveNext
loop
call closeRS(rs)

'Close Database
call closedb()

'If the config file (or some of the fields in it) is empty (as would 
'be the case for users who created their own database), then 
'pre-populate some fields with default values.
if isNull(urlNonSSL) or isEmpty(urlNonSSL) then
	urlNonSSL = "http://localhost/CandyPress/Scripts/"
end if
if isNull(urlSSL) or isEmpty(urlSSL) then
	urlSSL = "https://localhost/CandyPress/Scripts/"
end if
if isNull(pDownloadDir) or isEmpty(pDownloadDir) then
	pDownloadDir = "../Downloads/"
end if
if isNull(pImagesDir) or isEmpty(pImagesDir) then
	pImagesDir = "../ProdImages/"
end if
if isNull(pMaxCartQty) or isEmpty(pMaxCartQty) then
	pMaxCartQty = 30
end if
if isNull(pMaxItemQty) or isEmpty(pMaxItemQty) then
	pMaxItemQty = 20
end if
if isNull(pMinCartAmount) or isEmpty(pMinCartAmount) then
	pMinCartAmount = 0
end if
if isNull(pMaxItemsPerPage) or isEmpty(pMaxItemsPerPage) then
	pMaxItemsPerPage = 5
end if
if isNull(pOrderPrefix) or isEmpty(pOrderPrefix) then
	pOrderPrefix = "CP"
end if
if isNull(pCurrencySign) or isEmpty(pCurrencySign) then
	pCurrencySign = "$"
end if
if isNull(pStoreLCID) or isEmpty(pStoreLCID) then
	pStoreLCID = "1033"
end if
if isNull(pMaxDownloadHours) or isEmpty(pMaxDownloadHours) then
	pMaxDownloadHours = 0
end if
if isNull(pMaxDownloadCount) or isEmpty(pMaxDownloadCount) then
	pMaxDownloadCount = 0
end if
if isNull(pHideAddStockLevel) or isEmpty(pHideAddStockLevel) then
	pHideAddStockLevel = -1
end if
if isNull(handlingFeeAmt) or isEmpty(handlingFeeAmt) then
	handlingFeeAmt = 0
end if
if isNull(payMsgMailIn) or isEmpty(payMsgMailIn) then
	payMsgMailIn = "Mail-In"
end if
if isNull(payMsgCallIn) or isEmpty(payMsgCallIn) then
	payMsgCallIn = "Call-In"
end if
if isNull(payMsgFaxIn) or isEmpty(payMsgFaxIn) then
	payMsgFaxIn = "Fax-In"
end if
if isNull(payMsgCOD) or isEmpty(payMsgCOD) then
	payMsgCOD = "COD"
end if
if isNull(payMsgCreditCard) or isEmpty(payMsgCreditCard) then
	payMsgCreditCard = "CreditCard"
end if
if isNull(payMsgPayPal) or isEmpty(payMsgPayPal) then
	payMsgPayPal = "PayPal"
end if
if isNull(payMsgTwoCheckOut) or isEmpty(payMsgTwoCheckOut) then
	payMsgTwoCheckOut = "2CheckOut"
end if
if isNull(payMsgAuthNet) or isEmpty(payMsgAuthNet) then
	payMsgAuthNet = "Authorize.Net"
end if
if isNull(payMsgCustom) or isEmpty(payMsgCustom) then
	payMsgCustom = "Custom Payment"
end if
if isNull(payMsgOther) or isEmpty(payMsgOther) then
	payMsgOther = "Undetermined"
end if
if isNull(payMsgNotReq) or isEmpty(payMsgNotReq) then
	payMsgNotReq = "Payment Not Required"
end if
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Store Configuration</font></b>
</P>

<SCRIPT language="JavaScript">
<!--
	function openPopup(helpField) 
	{
		var w       = 450;
		var h       = 250;
		var popAttr = "width="+w+",height="+h+",resizable=1,scrollbars=1";
		var destURL = "utilities_ConfigHelp.asp?helpField="+helpField;
		window.open(destURL,"Help",popAttr);
	}
//-->
</SCRIPT>

<form method="post" action="utilities_configexec.asp" name="configMod">

<!-- ************************************************************** -->

<span class="textBlockHead">URL's and Folders</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<%call dispTextConfig("Full NON-SSL URL to ""Scripts"" folder&nbsp;(<a href=""javascript: void(0)"" target=""_blank"" onclick=""window.open(document.configMod.urlNonSSL.value);return false;"">Test</a>)",30,250,"urlNonSSL",urlNonSSL)%>
<%call dispTextConfig("Full SSL URL to ""Scripts"" folder&nbsp;(<a href=""javascript: void(0)"" target=""_blank"" onclick=""window.open(document.configMod.urlSSL.value);return false;"">Test</a>)",30,250,"urlSSL",urlSSL)%>
<%call dispTextConfig("Download folder (relative to /Scripts folder)",30,250,"pDownloadDir",pDownloadDir)%>
<%call dispTextConfig("Images folder (relative to /Scripts folder)",30,250,"pImagesDir",pImagesDir)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">Email</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" width="340">
		Mail component you will be using
	</td>
	<td bgcolor="#EEEEEE">
		<select name=mailComp id=mailComp size=1>
			<option value="0" <%=checkMatch(mailComp,"0")%>>None</option>
			<option value="1" <%=checkMatch(mailComp,"1")%>>JMail</option>
			<option value="2" <%=checkMatch(mailComp,"2")%>>CDONTS</option>
			<option value="3" <%=checkMatch(mailComp,"3")%>>Persits ASPemail</option>
			<option value="4" <%=checkMatch(mailComp,"4")%>>ServerObjects ASPmail</option>
			<option value="5" <%=checkMatch(mailComp,"5")%>>Bamboo SMTP</option>
			<option value="6" <%=checkMatch(mailComp,"6")%>>CDOSYS</option>
		</select>
	</td>
	<td bgcolor="#DDDDDD" align="center" valign="top" width="16">
		<a href="javascript:openPopup('mailComp');"><img src="x_help.gif" border="0"></a>
	</td>
</tr>
<%call dispTextConfig("Mail server's address",30,100,"pSMTPServer",pSMTPServer)%>
<%call dispTextConfig("Email address of Sales Department",30,100,"pEmailSales",pEmailSales)%>
<%call dispTextConfig("Email address of Webmaster or Technical contact",30,100,"pEmailAdmin",pEmailAdmin)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">Company Info</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<%call dispTextConfig("Company Name",30,100,"pCompany",pCompany)%>
<tr>
	<td bgcolor="#EEEEEE" valign=top>
		Company Address info (max 250 chars)
	</td>
	<td bgcolor="#EEEEEE" valign=top>
		<textarea name="pCompanyAddr" cols="30" rows="4" wrap=off><%=pCompanyAddr%></textarea>
	</td>
	<td bgcolor="#DDDDDD" align="center" valign="top">
		<a href="javascript:openPopup('pCompanyAddr');"><img src="x_help.gif" border="0"></a>
	</td>
</tr>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">General Settings</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<%call dispYNConfig("Operate store in Catalog Only mode?","pCatalogOnly",pCatalogOnly)%>
<tr>
	<td bgcolor="#EEEEEE" width="340">
		Select product details page layout
	</td>
	<td bgcolor="#EEEEEE">
		<select name=prodViewLayout id=prodViewLayout size=1>
			<option value="0" <%=checkMatch(prodViewLayout,"0")%>>Classic</option>
			<option value="1" <%=checkMatch(prodViewLayout,"1")%>>Professional</option>
		</select>
	</td>
	<td bgcolor="#DDDDDD" align="center" valign="top" width="16">
		<a href="javascript:openPopup('prodViewLayout');"><img src="x_help.gif" border="0"></a>
	</td>
</tr>
<tr>
	<td bgcolor="#EEEEEE" width="340">
		Select product list page layout
	</td>
	<td bgcolor="#EEEEEE">
		<select name=listViewLayout id=listViewLayout size=1>
			<option value="0" <%=checkMatch(listViewLayout,"0")%>>Classic</option>
			<option value="1" <%=checkMatch(listViewLayout,"1")%>>Extended</option>
		</select>
	</td>
	<td bgcolor="#DDDDDD" align="center" valign="top" width="16">
		<a href="javascript:openPopup('listViewLayout');"><img src="x_help.gif" border="0"></a>
	</td>
</tr>
<tr>
	<td bgcolor="#EEEEEE">
		Calculate taxes based on Billing or Shipping location?
	</td>
	<td bgcolor="#EEEEEE">
		<select name=taxBillOrShip id=taxBillOrShip size=1>
			<option value="0" <%=checkMatch(taxBillOrShip,"0")%>>Billing</option>
			<option value="1" <%=checkMatch(taxBillOrShip,"1")%>>Shipping</option>
		</select>
	</td>
	<td bgcolor="#DDDDDD" align="center" valign="top">
		<a href="javascript:openPopup('taxBillOrShip');"><img src="x_help.gif" border="0"></a>
	</td>
</tr>
<%call dispYNConfig("Hide pricing if Product Price is 0.00?","pHidePricingZero",pHidePricingZero)%>
<%call dispTextConfig("Maximum quantity allowed for the entire order",10,10,"pMaxCartQty",pMaxCartQty)%>
<%call dispTextConfig("Maximum quantity allowed per item",10,10,"pMaxItemQty",pMaxItemQty)%>
<%call dispTextConfig("Minimum purchase amount before checkout",10,10,"pMinCartAmount",pMinCartAmount)%>
<%call dispTextConfig("Number of products per page on product list page",10,10,"pMaxItemsPerPage",pMaxItemsPerPage)%>
<%call dispTextConfig("Order Number prefix",10,10,"pOrderPrefix",pOrderPrefix)%>
<%call dispTextConfig("Currency sign for your store",10,10,"pCurrencySign",pCurrencySign)%>
<%call dispTextConfig("Locale Identifier used to format dates & numbers",10,10,"pStoreLCID",pStoreLCID)%>
<%call dispTextConfig("Default Country Code (leave empty for none)",2,2,"defaultCountryCode",defaultCountryCode)%>
<%call dispYNConfig("Prevent modification of ""Email To Friend"" message body?","pEmailFriendSec",pEmailFriendSec)%>
<%call dispTextConfig("Max hours allowed to download software (0=unlimited)",10,10,"pMaxDownloadHours",pMaxDownloadHours)%>
<%call dispTextConfig("Max times allowed to download software (0=unlimited)",10,10,"pMaxDownloadCount",pMaxDownloadCount)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">Stock</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<%call dispYNConfig("Show ""In Stock"", ""Out of Stock"" messages?","pShowStockView",pShowStockView)%>
<%call dispYNConfig("Update stock if order status is Pending?","statUpdPending",statUpdPending)%>
<%call dispTextConfig("Out of stock level (-1 disables)",10,10,"pHideAddStockLevel",pHideAddStockLevel)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">Shipping</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" width="340">
		Select shipping rates display format
	</td>
	<td bgcolor="#EEEEEE">
		<select name=shipDisplayType id=shipDisplayType size=1>
			<option value="0" <%=checkMatch(shipDisplayType,"0")%>>List Box</option>
			<option value="1" <%=checkMatch(shipDisplayType,"1")%>>Radio Buttons</option>
		</select>
	</td>
	<td bgcolor="#DDDDDD" align="center" valign="top" width="16">
		<a href="javascript:openPopup('shipDisplayType');"><img src="x_help.gif" border="0"></a>
	</td>
</tr>
<%call dispYNConfig("Allow customer to enter a separate shipping address?","allowShipAddr",allowShipAddr)%>
<%call dispYNConfig("Include shipping total when calculating taxes?","taxOnShipping",taxOnShipping)%>
<%call dispTextConfig("Handling fee amount (0 for none)",10,10,"handlingFeeAmt",handlingFeeAmt)%>
<%call dispYNConfig("Include handling fee when calculating taxes?","handlingFeeTax",handlingFeeTax)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">Offline Payments</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" width="340">
		<b>1. Mail In Payments :</b>
	</td>
	<td bgcolor="#EEEEEE">&nbsp;</td>
	<td bgcolor="#DDDDDD" width="16">&nbsp;</td>
</tr>
<%call dispYNConfig("Allow payments to be Mailed in?","pMailIn",pMailIn)%>
<%call dispTextConfig("Mail In payments description",30,50,"payMsgMailIn",payMsgMailIn)%>
<tr>
	<td bgcolor="#EEEEEE" width="340">
		<b>2. Call In Payments :</b>
	</td>
	<td bgcolor="#EEEEEE">&nbsp;</td>
	<td bgcolor="#DDDDDD" width="16">&nbsp;</td>
</tr>
<%call dispYNConfig("Allow payments to be Called in?","payCallIn",payCallIn)%>
<%call dispTextConfig("Call In payments description",30,50,"payMsgCallIn",payMsgCallIn)%>
<tr>
	<td bgcolor="#EEEEEE" width="340">
		<b>3. Fax In Payments :</b>
	</td>
	<td bgcolor="#EEEEEE">&nbsp;</td>
	<td bgcolor="#DDDDDD" width="16">&nbsp;</td>
</tr>
<%call dispYNConfig("Allow payments to be Faxed in?","payFaxIn",payFaxIn)%>
<%call dispTextConfig("Fax In payments description",30,50,"payMsgFaxIn",payMsgFaxIn)%>
<tr>
	<td bgcolor="#EEEEEE" width="340">
		<b>4. COD Payments :</b>
	</td>
	<td bgcolor="#EEEEEE">&nbsp;</td>
	<td bgcolor="#DDDDDD" width="16">&nbsp;</td>
</tr>
<%call dispYNConfig("Allow COD payments?","payCOD",payCOD)%>
<%call dispTextConfig("COD payments description",30,50,"payMsgCOD",payMsgCOD)%>
<tr>
	<td bgcolor="#EEEEEE" width="340">
		<b>5. Offline Credit Card Payments :</b>
	</td>
	<td bgcolor="#EEEEEE">&nbsp;</td>
	<td bgcolor="#DDDDDD" width="16">&nbsp;</td>
</tr>
<%call dispYNConfig("Allow Offline Credit Card payments?","pCreditCard",pCreditCard)%>
<%call dispTextConfig("Credit Card payments description",30,50,"payMsgCreditCard",payMsgCreditCard)%>
<%call dispTextConfig("List of accepted credit cards (seperated by commas)",30,250,"pCCType",pCCType)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">PayPal Payments</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<%call dispYNConfig("Allow payments via PayPal?","pPayPal",pPayPal)%>
<%call dispTextConfig("PayPal payments description",30,50,"payMsgPayPal",payMsgPayPal)%>
<%call dispTextConfig("Your PayPal member ID",30,100,"payPalMemberID",payPalMemberID)%>
<tr>
	<td bgcolor="#EEEEEE" width="340">
		Currency to use for PayPal (Subject to availability)
	</td>
	<td bgcolor="#EEEEEE">
		<select name=payPalCurrCode id=payPalCurrCode size=1>
			<option value="USD" <%=checkMatch(payPalCurrCode,"USD")%>>US Dollar</option>
			<option value="CAD" <%=checkMatch(payPalCurrCode,"CAD")%>>Canadian Dollar</option>
			<option value="EUR" <%=checkMatch(payPalCurrCode,"EUR")%>>Euro</option>
			<option value="GBP" <%=checkMatch(payPalCurrCode,"GBP")%>>Pounds Sterling</option>
			<option value="JPY" <%=checkMatch(payPalCurrCode,"JPY")%>>Japanese Yen</option>
		</select>
	</td>
	<td bgcolor="#DDDDDD" align="center" valign="top" width="16">
		<a href="javascript:openPopup('payPalCurrCode');"><img src="x_help.gif" border="0"></a>
	</td>
</tr>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">2CheckOut Payments</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<%call dispYNConfig("Allow payments via 2CheckOut?","TwoCheckOut",TwoCheckOut)%>
<%call dispTextConfig("2CheckOut payments description",30,50,"payMsgTwoCheckOut",payMsgTwoCheckOut)%>
<%call dispTextConfig("Your 2CheckOut Account Number",30,100,"TwoCheckOutSID",TwoCheckOutSID)%>
<%call dispTextConfig("Your 2CheckOut MD5 Secret Word",30,100,"TwoCheckOutMD5",TwoCheckOutMD5)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">Authorize.Net Payments</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<%call dispYNConfig("Allow payments via Authorize.Net?","pAuthNetFrontEnd",pAuthNetFrontEnd)%>
<%call dispYNConfig("Allow admin authorizations via Authorize.Net?","pAuthNet",pAuthNet)%>
<%call dispTextConfig("Authorize.Net payments description",30,50,"payMsgAuthNet",payMsgAuthNet)%>
<%call dispTextConfig("Authorize.Net Login ID",30,100,"authNetLogin",authNetLogin)%>
<%call dispTextConfig("Authorize.Net Transaction Key",30,100,"authNetTxKey",authNetTxKey)%>
<%call dispTextConfig("Authorize.Net Currency Code for your store",5,10,"authNetCurrCode",authNetCurrCode)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">Custom Payments</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<%call dispYNConfig("Use custom payment routine?","payCustom",payCustom)%>
<%call dispTextConfig("Custom payments description",30,50,"payMsgCustom",payMsgCustom)%>
</table>

<br>

<!-- ************************************************************** -->

<span class="textBlockHead">Miscellaneous Payment Info</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" width="340">
		Default Payment Type
	</td>
	<td bgcolor="#EEEEEE">
		<select name=payDefault id=payDefault size=1>
			<option value=""             <%=checkMatch(payDefault,"")            %>>None</option>
			<option value="MailIn"       <%=checkMatch(payDefault,"MailIn")      %>>MailIn</option>
			<option value="PayPal"       <%=checkMatch(payDefault,"PayPal")      %>>PayPal</option>
			<option value="2CheckOut"    <%=checkMatch(payDefault,"2CheckOut")   %>>2CheckOut</option>
			<option value="AuthorizeNet" <%=checkMatch(payDefault,"AuthorizeNet")%>>AuthorizeNet</option>
			<option value="CreditCard"   <%=checkMatch(payDefault,"CreditCard")  %>>CreditCard</option>
			<option value="Custom"       <%=checkMatch(payDefault,"Custom")      %>>Custom</option>
		</select>
	</td>
	<td bgcolor="#DDDDDD" align="center" valign="top" width="16">
		<a href="javascript:openPopup('payDefault');"><img src="x_help.gif" border="0"></a>
	</td>
</tr>
<%call dispTextConfig("Other/Unknown payments description",30,50,"payMsgOther",payMsgOther)%>
<%call dispTextConfig("Description to be used if payment is NOT required",30,50,"payMsgNotReq",payMsgNotReq)%>
</table>

<br>

<center>
	<input type="submit" name="submit1" value="Update Configuration">
</center>

</form>

<!--#include file="_INCfooter_.asp"-->

<%
'**********************************************************************
'Display Text Configuration Settings
'**********************************************************************
sub dispTextConfig(Description,Size,MaxLength,Name,Value)
%>
	<tr>
		<td bgcolor="#EEEEEE" width="340"><%=Description%></td>
		<td bgcolor="#EEEEEE">
			<input type="text" size="<%=Size%>" maxlength="<%=MaxLength%>" name="<%=Name%>" value="<%=Value%>"> 
		</td>
		<td bgcolor="#DDDDDD" align="center" valign="top" width="16">
			<a href="javascript:openPopup('<%=Name%>');"><img src="x_help.gif" border="0"></a>
		</td>
	</tr>
<%
end sub
'**********************************************************************
'Display Y/N Configuration Settings
'**********************************************************************
sub dispYNConfig(Description,Name,Value)
%>
	<tr>
		<td bgcolor="#EEEEEE" width="340"><%=Description%></td>
		<td bgcolor="#EEEEEE">
			<select name="<%=name%>" id="<%=name%>" size="1">
				<option value="0"  <%=checkMatch(Value,"0") %>>No</option>
				<option value="-1" <%=checkMatch(Value,"-1")%>>Yes</option>
			</select>
		</td>
		<td bgcolor="#DDDDDD" align="center" valign="top" width="16">
			<a href="javascript:openPopup('<%=name%>');"><img src="x_help.gif" border="0"></a>
		</td>
	</tr>
<%
end sub
%>
