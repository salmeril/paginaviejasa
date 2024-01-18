<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Modify Store General Settings - Help
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

'Variables
dim helpField
'*************************************************************************
%>
<html>
<head>
	<title>Store Configuration - Help</title>
	<STYLE type="text/css">
	<!--
	BODY   {COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt;}
	B      {COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt; FONT-WEIGHT: bold}
	TD     {COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt}
	P      {COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt}
	.examp {COLOR: #333333; FONT-FAMILY: Courier, Verdana, Arial, helvetica; FONT-SIZE: 8pt}
	-->
	</STYLE>
</head>

<body>

<p align=center>
	<b><font size=2>Store Configuration - Help</font></b>
</p>

<a href="javascript:parent.self.close()">Close Window</a>
<hr>

<%
'Determine which field we are showing help for
helpField = trim(Request.QueryString("helpField"))

'Display Help Field Name
Response.Write "<b><font color=#800000>NAME:</font></b> " & helpField & "<br><br>"
Response.Write "<b><font color=#800000>DESCRIPTION:</font></b> "

'Display Help text
select case lCase(helpField)
case "urlnonssl"
%>
	Full <b>non-SSL</b> (ie. non-secure) URL to your store's "scripts" 
	folder. This value should start with the characters "http://" 
	and end with "/scripts/".
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">http://www.myServer.com/CandyPress/scripts/</span>
<%
case "urlssl"
%>	  
	Full <b>SSL</b> (ie. secure) URL to your store's "scripts" folder. 
	Typically, this value would start with the characters "https://" 
	and end with "/scripts/".
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">https://www.myServer.com/CandyPress/scripts/</span>
	<br><br>
	<b>Note :</b> If you simply want to test the 
	store, or you don't have SSL yet, you can enter the same 
	value you entered into the non-SSL field. The store would 
	still be fully functional, but the checkout process will not 
	be secure. 
<%
case "pdownloaddir"
%>	  
	Location of you Software Download Directory RELATIVE to the 
	"scripts" folder. If your store will sell items that are 
	downloadable from your web site, you must enter the path to the 
	folder in which the downloadable items will be stored into this 
	field. If you don't plan on selling downloadable items, you can 
	simply set this value to "../". The location must start with "../" 
	and end with "/".
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">../Downloads/</span>
<%
case "pimagesdir"
%>	  
	Location of your Product Images Directory RELATIVE to the "scripts" 
	folder. The location must start with "../" and end with "/".
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">../ProdImages/</span>
<%
case "mailcomp"
%>	  
	This tells the software what type of email component you will be 
	using to send email from your store. Check with your web hosting 
	company if you don't know what email component(s) are available 
	to you. We provide for several of the most popular components.
<%
case "psmtpserver"
%>	  
	This is the address of your SMTP server.
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">mail.myServer.com</span>
<%
case "pemailsales"
%>	  
	This is the email address of the person or department that will 
	be responsible for sales related queries.
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">sales@myServer.com</span>
<%
case "pemailadmin"
%>	  
	This is the email address of the 
	person or department which will be responsible for technical issues 
	related to your store.
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">webmaster@myServer.com</span>
<%
case "pcompany"
%>	  
	The name of your store or company. 
	This value is diplayed in various places in your store.
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">Madison Shoe Store Inc.</span>
<%
case "pcompanyaddr"
%>	  
	Your company's address. 
	Maximum number of characters allowed is 250. Enter address info 
	on seperate lines in the text box. Try to limit the number of 
	address lines to no more than 5 or 6 to ensure proper display 
	on invoices etc.
<%
case "pcatalogonly"
%>	  
	Sets the store to run in "Catalog" mode. 
	This means that the customer will be able to browse 
	your store, but won't be able to add any of the products to a 
	shopping cart.
<%
case "prodviewlayout"
%>	  
	Selects the type of layout to use when displaying the product's 
	details.
<%
case "listviewlayout"
%>	  
	Selects the type of layout to use on the product list pages, 
	such as the home page, search results, category list, etc. The 
	"Extended" layout is the same as the "Classic" layout, except for 
	the additional display of Stock Levels, Ratings, and Free Shipping 
	messages. In addition, the home page will also show the list price 
	and savings in "Extended" layout mode.
<%
case "shipdisplaytype"
%>	  
	Selects the type of layout to use when displaying shipping rates 
	for the order.
<%
case "allowshipaddr"
%>	  
	If set to Yes, the customer will be able to enter a shipping 
	address that is different from their billing address. If set to 
	No, the customer will not have the opportunity to enter a seperate 
	shipping address, and the system will assume that the shipping 
	address is the same as the billing address.
<%
case "taxonshipping"
%>	  
	Determines if the total shipping cost will be included when taxes 
	are calculated for the order.
<%
case "handlingfeeamt"
%>	  
	Use this setting to add a Handling Fee to orders. The handling 
	fee will apply to all orders where the total shipping weight is 
	greater than 0.
<%
case "handlingfeetax"
%>	  
	If a handling fee is charged, this setting will determine if the 
	fee is taxable or not.
<%
case "taxbillorship"
%>	  
	Determines if taxes for the order must be calculated based on the 
	Billing or Shipping address for the order. Note : If no shipping 
	address was entered, the billing and shipping address is the same.
<%
case "phidepricingzero"
%>	  
	Hides all pricing information 
	on Product List and Detail pages, but only if the price of the product 
	is 0.00 (zero), otherwise show prices. This setting is usefull if you 
	want to carry products where the entire product price is made up of 
	options. This way you can enter a product price of '0.00', and link 
	'required' options with a price value to that product.
<%
case "pshowstockview"
%>	  
	Configures the store to show or hide 
	the stock level messages (ie. "In Stock", "Out of Stock").
<%
case "phideaddstocklevel"
%>	  
	The product list and 
	product detail pages will check the product's stock level against 
	the value you enter into this field. If the product's stock level 
	is at or below the level you specify, the 'Add' buttons will 
	automatically be hidden, and an "Out of Stock" message will be 
	displayed. To disable stock levels checking, enter a "-1" into 
	this field. This will result in the product always being displayed 
	as "In Stock" regardless of actual stock levels.
<%
case "statupdpending"
%>
	Normally, an item's stock level is reduced when the order status 
	changes to "Paid", "Complete" or "Shipped". In addition to these, 
	you may specify if the item's stock level must be reduced when the 
	order status changes to "Pending". This normally happens when the 
	order is saved, and before payment is made.
<%
case "pmaxcartqty"
%>	  
	Maximum number of items (quantity) per Order.
<%
case "pmaxitemqty"
%>	  
	Maximum number of items per Product per Order.
<%
case "pmincartamount"
%>	  
	Minimum total purchase 
	price per Order BEFORE Tax and Shipping. Note : If you offer free 
	items (free downloads, brochures, etc.) there is a possibility that 
	you can have an order where the total is 0.00. Setting 
	the minimum order amount to a value greater than 0.00 will 
	disallow such orders.
<%
case "pmaxitemsperpage"
%>	  
	The maximum number of Items you want 
	to display, per page, on the Product Search/View pages.
<%
case "porderprefix"
%>	  
	Prefix to be displayed as part of the Order Number throughout 
	the store.
<%
case "pcurrencysign"
%>	  
	Currency symbol for your Store.
	<br><br>
	<b><font color=#800000>EXAMPLE:</font></b> 
	<span class="examp">$, CAN$, &pound;
<%
case "defaultcountrycode"
%>	  
	Default Country Code for the store. This setting is mainly used 
	on the Customer Information page during checkout. The country 
	code must correspond to the country code for that country in the 
	database.
<%
case "pstorelcid"
%>	  
	Locale Identifier (LCID) for the store. An LCID is a decimal 
	numeric value that tells the web server what format to use when it 
	works with dates, numbers, etc. For example, the default locale 
	used by this software for it's internal date functions is '1033' 
	(US English). This can not be changed. You can however, override 
	the external (displayed) date format. We listed a few LCID 
	examples below. The entire list can be found at 
	<a href=http://www.devguru.com/Technologies/vbscript/quickref/LCIDchart.html target="_blank">DevGuru.Com</a>.<br><br>
	<table border=1>
		<tr>
			<td><b>LCID</b></td>
			<td><b>Description</b></td>
			<td><b>Date</b></td>
			<td><b>Number</b></td>
		</tr>
		<tr>
			<td>3081</td>
			<td>English - Australia</td>
			<td>17/04/2002</td>
			<td>1,350.15</td>
		</tr>
		<tr>
			<td>4105</td>
			<td>English - Canada</td>
			<td>17/04/2002</td>
			<td>1,350.15</td>
		</tr>
		<tr>
			<td>2057</td>
			<td>English - United Kingdom</td>
			<td>17/04/2002</td>
			<td>1,350.15</td>
		</tr>
		<tr>
			<td>1033</td>
			<td>English - USA</td>
			<td>4/17/2002</td>
			<td>1,350.15</td>
		</tr>
		<tr>
			<td>1036</td>
			<td>French - Standard</td>
			<td>17/04/2002</td>
			<td>1 350,15</td>
		</tr>
		<tr>
			<td>1031</td>
			<td>German - Standard</td>
			<td>17.04.2002</td>
			<td>1.350,15</td>
		</tr>
		<tr>
			<td>1040</td>
			<td>Italian - Standard</td>
			<td>17/04/2002</td>
			<td>1.350,15</td>
		</tr>
		<tr>
			<td>1034</td>
			<td>Spanish - Standard</td>
			<td>17/04/2002</td>
			<td>1.350,15</td>
		</tr>
	</table>
<%
case "pemailfriendsec"
%>	  
	If you set this value to "Yes", the 
	customer will no longer be able to modify the "Email To Friend" 
	message body. If set to "No", the customer will be able to enter 
	their own message in the message body. The problem with allowing 
	anyone to enter their own message, is that they could possibly 
	use the script to send personal email messages to others 
	"anonymously". 
<%
case "pmaxdownloadhours"
%>	  
	This setting specifies the 
	maximum number of hours you will allow a customer to download 
	software from your store after the FIRST download has taken place. 
	Be carefull to allow enough time as some customers may need to try 
	several times due to bad connections, etc. Enter a "0" if you 
	want to disable this feature (ie. "0" allows for unlimited download 
	hours).
<%
case "pmaxdownloadcount"
%>	  
	This setting specifies the 
	maximum number of times you will allow a customer to download 
	software from your store. Enter a "0" if you want to disable this 
	feature (ie. "0" allows for unlimited download times).
<%
case "pmailin"
%>	  
	Allow/Disallow Mail-In payments such as 
	Checks, Money Orders, Cash, etc.
<%
case "paycallin"
%>	  
	Allow/Disallow Call-In payments where 
	the customers phones you, or you phone the 
	customer to gather payment information.
<%
case "payfaxin"
%>	  
	Allow/Disallow Fax-In payments where the 
	customer faxes their payment information to you.
<%
case "paycod"
%>	  
	Allow/Disallow COD payments.
<%
case "ppaypal"
%>	  
	Allow/Disallow payments via PayPal. 
	If enabled, the customer would be directed to PayPal's web site 
	for payment.
<%
case "paypalmemberid"
%>	  
	If you accept PayPal payments 
	on your site, specify your PayPal Membership ID (this is usually in 
	the form of an email address).
<%
case "paypalcurrcode"
%>	  
	PayPal accepts payments in several different currencies. Your 
	products should be priced in the same currency as the one 
	you select here. In other words, if you select US Dollar as 
	your PayPal currency, your products should also be priced in 
	US Dollars in your store.<br><br>
	Note : This feature will only be available during mid November, 
	2002. Please check the PayPal web site to ensure that this 
	feature is available before you use it. Until PayPal implements 
	the feature, all payments will be in US Dollars only.
<%
case "twocheckout"
%>	  
	Allow/Disallow payments via 2CheckOut. 
	If enabled, the customer will be directed to 2CheckOut.Com's web 
	site for payment.
<%
case "twocheckoutsid"
%>	  
	If you accept 2CheckOut payments 
	on your site, specify your 2CheckOut Account Number.
<%
case "twocheckoutmd5"
%>	  
	If you accept 2CheckOut payments 
	on your site, and want to automatically update the order's status, 
	then you must enter your MD5 "Secret Word" here. Please 
	note that this is NOT the same as the password that you use to 
	logon to 2CheckOut.Com's website. At the time of writing, this 
	value is found and modified by logging on to your 2CheckOut.Com 
	account, and clicking on "Account Details -&gt; Return". At the 
	very bottom of the page, you will see a box named "Secret Word". 
	This is the value you must enter.
<%
case "pcreditcard"
%>	  
	Allow/Disallow Credit Card payments. 
	If enabled, the customer will enter their Credit Card information 
	in a form on your store's web site, and the information will be 
	saved in your database for processing later on. Note 
	that you would typically need to have some form of Merchant Account 
	with a bank to use this form of payment. See below for more settings 
	related to credit card payments.
<%
case "pcctype"
%>	  
	If you are able to 
	directly accept and process transactions for Credit Cards, enter 
	the list of Credit Cards you are able to accept, seperated by commas, 
	into this field.
	<br><br>
	Eg. "Visa,MasterCard,American Express"
<%
case "pauthnetfrontend"
%>	  
	Allow/Disallow Authorize.Net payments. 
	If enabled, the customer will be directed to Authorize.Net's web 
	site for payment.
<%
case "pauthnet"
%>	  
	If you use offline Credit Card payments, and have an Authorize.Net 
	account, you can use the special Authorize.Net admin function to 
	assist in the authorization process. This saves you from manually 
	entering the offline credit card info into your Authorize.Net 
	virtual terminal screen.
	<br><br>
	Note : If you use offline Credit Card payments, but you DON'T have 
	an Authorize.Net account, you must use the virtual terminal provided 
	by your merchant account provider.
<%
case "authnetlogin"
%>	  
	The Login given to you when you opened your Authorize.Net account.
<%
case "authnettxkey"
%>	  
	The transaction key is used to create a secure fingerprint for each 
	transaction processed through Authorize.Net.
<%
case "authnetcurrcode"
%>	  
	The Authorize.Net Currency Code in which your Credit Card 
	transactions must settle (eg. "USD"). Contact Authorize.Net 
	for assistance if you are not sure what the Currency Code for 
	your account is.
<%
case "paydefault"
%>	  
	This will be the default payment type 
	when the customer checks out (ie. it will be pre-selected in the 
	drop down list of payment types available at your store). 
	Naturally, the default payment type should be one that you 
	support in your store.
<%
case "paymsgmailin"
%>	  
	The description that will be displayed 
	for "Mail In" payment types (eg. "Mail-In").
<%
case "paymsgcallin"
%>	  
	The description that will be displayed 
	for "Call In" payment types (eg. "Call-In").
<%
case "paymsgfaxin"
%>	  
	The description that will be displayed 
	for "Fax In" payment types (eg. "Fax-In").
<%
case "paymsgcod"
%>	  
	The description that will be displayed 
	for "COD" payment types (eg. "COD").
<%
case "paymsgcreditcard"
%>	  
	The description that will be displayed 
	for "Credit Card" payment types (eg. "Credit Card").
<%
case "paymsgpaypal"
%>	  
	The description that will be displayed 
	for "PayPal" payment types (eg. "PayPal").
<%
case "paymsgtwocheckout"
%>	  
	The description that will be displayed 
	for "2CheckOut" payment types (eg. "2CheckOut").
<%
case "paymsgauthnet"
%>	  
	The description that will be displayed 
	for "AuthorizeNet" payment types (eg. "Authorize.Net").
<%
case "paymsgother"
%>	  
	The description that will be displayed 
	when a Payment Type could not be determined or if a Payment Type was 
	not entered (eg. "Unknown").
<%
case "paymsgnotreq"
%>	  
	This special Payment Type description 
	deals with Orders that are free to the customer (eg. "Payment not 
	Required"). This could happen, for instance, when you provide for 
	free downloads of software items. It is likely that you would then 
	have an order where the total is 0.00. 
<%
case "paycustom"
%>	  
	Allow/Disallow custom payment routine. 
	Before you enable this payment method, you will have 
	to add the necessary code in the custom payment user exit files 
	to send the appropriate order information to the payment gateway, 
	and receive the payment status from the gateway. See the 
	"readme.htm" file, and the notes and examples inside 
	"_INCpayOut_.asp" and "_INCpayIn_.asp" for more details. 
	Some ASP knowledge will be required.
<%
case "paymsgcustom"
%>	  
	The description that will be displayed 
	for "Custom" payment types (eg. "Custom Payment").
<%
case else
%>
	<br><br><br>
	<b>Invalid help field was passed...</b>
	<br><br><br>
<%
end select
%>

<hr>
<a href="javascript:parent.self.close()">Close Window</a>

</body>

</html>
