<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Authorize.Net payment form
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
<!--#include file="../Scripts/_INCauthNet_.asp"-->
<%
'Work Fields
dim mySQL, cn, rs
dim authNetDemo

'cartHead
dim idOrder
dim idCust
dim orderDate
dim Total
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
dim paymentType
dim cardType
dim cardNumber
dim cardExpMonth
dim cardExpYear
dim cardVerify
dim cardName
dim orderStatus

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Get idOrder
idOrder = trim(Request.QueryString("recId"))
if len(idOrder) = 0 then
	idOrder = trim(Request.Form("recId"))
end if
if idOrder = "" or not isNumeric(idOrder) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
end if

'Are we in demo mode?
if demoMode = "Y" then
	authNetDemo = "TRUE"
else
	authNetDemo = "FALSE"
end if

%>
<!--#include file="_INCheader_.asp"-->

<SCRIPT LANGUAGE="JavaScript">
<!--
function popW() {return 500;}
function popH() {return 400;}
function popL() {return (screen.width - popW()) / 2;}
function popT() {return (screen.height - popH()) / 2;}
-->
</script>

<P align=left>
	<b><font size=3>Authorize.Net Payment Form</font></b>
	<br><br>
</P>
<%

'Get cartHead Record
mySQL	= "SELECT idCust,orderDate,total,name,lastName,customerCompany," _
		& "       phone,email,address,city,locState,locCountry,zip," _
		& "       paymentType,cardType,cardNumber,cardExpMonth," _
		& "       cardExpYear,cardVerify,cardName,orderStatus " _
		& "FROM   cartHead " _
		& "WHERE  idOrder=" & idOrder
set rs = openRSexecute(mySQL)
if rs.eof then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
else
	idCust  			= rs("idCust")
	orderDate			= rs("orderDate")
	Total				= rs("Total")
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
	paymentType			= trim(rs("paymentType"))
	cardType			= trim(rs("cardType"))
	cardNumber			= trim(EnDeCrypt(Hex2Ascii(rs("cardNumber")),rc4Key))
	cardExpMonth		= trim(rs("cardExpMonth"))
	cardExpYear			= trim(rs("cardExpYear"))
	cardVerify			= trim(rs("cardVerify"))
	cardName			= trim(rs("cardName"))
	orderStatus			= trim(rs("orderStatus"))
end if
call closeRS(rs)
%>
<span class="textBlockHead">Authorize Payment</span>
&nbsp;<%call maintNavLinks()%><br><br>

<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
<form action="https://secure.authorize.net/gateway/transact.dll" method="post" name="authNet" id="authNet" onSubmit="window.open('',this.target,'width='+popW()+',height='+popH()+',top='+popT()+',left='+popL()+',resizable=yes,scrollbars=yes'); return true;" target="authNetWin">
	<%call InsertFP(authNetLogin,authNetTxKey,moneyD(total),idOrder,authNetCurrCode)%>
	<input type="hidden" name="x_version"				value="3.1">
	<input type="hidden" name="x_type"					value="AUTH_CAPTURE">
	<input type="hidden" name="x_Show_Form"				value="PAYMENT_FORM">
	<input type="hidden" name="x_method"				value="CC">
	<input type="hidden" name="x_Email_Customer"		value="TRUE">
	<input type="hidden" name="x_Email_Merchant"		value="TRUE">
	<input type="hidden" name="x_Amount"				value="<%=moneyD(total)%>">
	<input type="hidden" name="x_Email"					value="<%=email%>">
	<input type="hidden" name="x_cust_id"				value="<%=idCust%>">
	<input type="hidden" name="x_Description"			value="<%=pCompany & " " & pOrderPrefix & "-" & idOrder%>">
	<input type="hidden" name="x_Invoice_Num"			value="<%=pOrderPrefix & "-" & idOrder%>">
	<input type="hidden" name="x_currency_code"			value="<%=authNetCurrCode%>">
	<input type="hidden" name="x_Login"					value="<%=authNetLogin%>">
	<input type="hidden" name="x_Test_Request"			value="<%=authNetDemo%>">
	<!--
	<input type="hidden" name="x_password"				value="testdriver">
	<input type="hidden" name="x_Receipt_Link_URL"		value="https://www.server.com/Admin/xxx.asp">
	<input type="hidden" name="x_Receipt_Link_Method"	value="Post">
	<input type="hidden" name="x_Receipt_Link_Text"		value="Continue ...">
	-->
	<TR>
		<TD align="left">
			<table border="0">
				<tr>
					<td valign=top nowrap><b>Order Number</b>&nbsp;</td>
					<td valign=top><%=pOrderPrefix & "-" & idOrder%></td>
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
					<td valign=top nowrap><b>Payment Type</b>&nbsp;</td>
					<td valign=top><%=paymentType%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Card Type</b>&nbsp;</td>
					<td valign=top><%=cardType%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Card Name</b>&nbsp;</td>
					<td valign=top><%=cardName%></td>
				</tr>
				<tr>
					<td valign=top nowrap><b>Amount</b>&nbsp;</td>
					<td valign=top><b><%=moneyD(total) & " (" & authNetCurrCode & ")"%></b></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>Card Number</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=18 MAXLENGTH=22 NAME="x_card_num" VALUE="<%=cardNumber%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>Card Expire</b>&nbsp;</td>
<%
					if isEmpty(cardExpMonth) or isnull(cardExpMonth) then
%>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=4 MAXLENGTH=4 NAME="x_exp_date" VALUE="">&nbsp;&nbsp;(MMYY)</td>
<%
					else
%>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=4 MAXLENGTH=4 NAME="x_exp_date" VALUE="<%=Replace((Space(2-Len(cardExpMonth))&cardExpMonth)," ","0") & right(cardExpYear,2)%>">&nbsp;&nbsp;(MMYY)</td>
<%
					end if
%>
				</tr>
				<tr>
					<td valign=middle nowrap><b>First Name</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=20 MAXLENGTH=50 NAME="x_first_name" VALUE="<%=Name%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>Last Name</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=20 MAXLENGTH=50 NAME="x_last_name" VALUE="<%=LastName%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>Address</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=20 MAXLENGTH=60 NAME="x_address" VALUE="<%=address%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>City</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=20 MAXLENGTH=40 NAME="x_city" VALUE="<%=city%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>Zip/PCode</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=10 MAXLENGTH=20 NAME="x_zip" VALUE="<%=zip%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>State/Province</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=20 MAXLENGTH=40 NAME="x_state" VALUE="<%=locState%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>Country</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=20 MAXLENGTH=60 NAME="x_country" VALUE="<%=locCountry%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>Company</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=20 MAXLENGTH=50 NAME="x_company" VALUE="<%=customerCompany%>"></td>
				</tr>
				<tr>
					<td valign=middle nowrap><b>Phone</b>&nbsp;</td>
					<td valign=middle><INPUT TYPE="TEXT" SIZE=20 MAXLENGTH=25 NAME="x_phone" VALUE="<%=phone%>"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align=center>
			<br>
			<input type="submit" name="submit" value="Authorize Payment">
			<br><br>
		</td>
	</tr>
</form>
</TABLE>

<br>
<span class="textBlockHead">Help and Instructions :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr><td>

	<b>General</b> - Some fields on this page is intentionally modifiable to 
	allow for maximum flexibility when authorizing Credit Card transactions. 
	In this way, it's possible to override the initial payment method (which 
	was selected by the Customer when they placed the order) and process 
	a Credit Card transaction instead (obviously you will have to contact 
	the customer for his/her credit card info before doing this). 
	Alternatively, you may find that a Credit Card transaction is rejected 
	by Authorize.Net due to incorrect Credit Card info. You could then 
	contact the customer and obtain the correct info before proceeding. 
	Note however that, should you modify any of these fields, the original 
	order will still remain the same. Only the information sent to 
	Authorize.Net will be different. If you want to modify these fields 
	in the order itself, you must use the "Edit Order" function.<br><br>
	
	<b>Currency Code</b> - If the Currency Code (eg. "USD) shown next to the 
	amount is incorrect, then you will need to enter the correct Currency 
	Code into your store's configurations. If you don't know what this 
	value should be, contact Authorize.Net for more information.<br><br>
	
	<b>Authorization Code</b> - Upon successful authorization of the 
	transaction, Authorize.Net will provide you with an Authorization Code 
	and some other usefull information such as the Authorization Date etc. 
	We advise you to copy this information into the "Private Comments" 
	field (Click on "Edit Order") so that you don't have to refer to 
	confirmation emails from Authorize.Net, or your Authorize.Net account, 
	for this this information.<br><br>
	
</td></tr>
</table>
<%
call closedb()
%>
<!--#include file="_INCfooter_.asp"-->
<%
'*********************************************************************
'Create Navigation Links
'*********************************************************************
sub maintNavLinks()
%>
	[ 
	<a href="SA_order.asp?recallCookie=1">List Orders</a> | 
	<a href="SA_order_edit.asp?action=edit&recid=<%=idOrder%>">Edit</a> | 
	<a href="SA_order_edit.asp?action=view&recid=<%=idOrder%>">View</a> | 
	<a href="SA_order_edit.asp?action=del&recid=<%=idOrder%>">Delete</a> 
	]
<%
end sub
%>
