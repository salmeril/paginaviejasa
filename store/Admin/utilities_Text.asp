<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Modify Store Text Configurations
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
<%
'Text Variables
dim termsAndCond
dim saveOrderEmail
dim paySuccessMsg
dim payErrorMsg
dim passRequestEmail
dim emailToFriend
dim statusUpdateEmail

'Database variables
dim mySQL, cn, rs

'Work Variables
dim tempConfigValLong
'*************************************************************************

'Open Database
call openDb()

'Get current text configurations from database
mySQL = "SELECT configVar, configValLong " _
	  & "FROM   storeAdmin " _
	  & "WHERE  adminType = 'T'"
set rs = openRSexecute(mySQL)
do while not rs.EOF

	'Get text
	tempConfigValLong =  rs("configValLong")

	'Assign text to local variable
	select case trim(lCase(rs("configVar")))
	case lCase("termsAndCond")
		termsAndCond			= tempConfigValLong
	case lCase("saveOrderEmail")
		saveOrderEmail			= tempConfigValLong
	case lCase("paySuccessMsg")
		paySuccessMsg			= tempConfigValLong
	case lCase("payErrorMsg")
		payErrorMsg				= tempConfigValLong
	case lCase("passRequestEmail")
		passRequestEmail		= tempConfigValLong
	case lCase("emailToFriend")
		emailToFriend			= tempConfigValLong
	case lCase("statusUpdateEmail")
		statusUpdateEmail		= tempConfigValLong
	end select

	rs.MoveNext
loop
call closeRS(rs)

'Close Database
call closedb()

'If some of the text configurations are empty (as would be the case 
'for users who created their own database), then pre-populate some 
'required fields with default values.

if isNull(termsAndCond) or isEmpty(termsAndCond) then
	termsAndCond = "" _
	& "<b>1. Payment</b> - State payment policy ..." & vbCrlf _
	& "<br><br>" & vbCrLf & vbCrLf _
	& "<b>2. Refunds</b> - State refund policy ..." & vbCrLf _
	& "<br><br>" & vbCrLf & vbCrLf _
	& "<b>3. Some more Rules</b> - Some more rules, etc.. " & vbCrlf _ 
	& "<br><br>"
end if
if isNull(saveOrderEmail) or isEmpty(saveOrderEmail) then
	saveOrderEmail = "" _
	& "#DATE#" & vbCrLf & vbCrLf _
	& "Dear #NAME#," & vbCrLf & vbCrLf _
	& "Thank you for ordering from our store. This is to confirm that we have received the following order." & vbCrLf & vbCrLf _
	& "Order Number : #ORDER#" & vbCrLf _
	& "Order Total  : #TOTAL#" & vbCrLf _
	& "Payment Type : #PAYMT#" & vbCrLf & vbCrLf _
	& "Please note that you can track your Order Status by logging on to your 'Account'." & vbCrLf & vbCrLf _
	& "If you have any questions regarding your order, please contact us via email at #SALES#."  & vbCrLf & vbCrLf _
	& "Regards" & vbCrLf & vbCrLf _
	& "#STORE#" & vbCrLf & vbCrLf
end if
if isNull(statusUpdateEmail) or isEmpty(statusUpdateEmail) then
	statusUpdateEmail = "" _
	& "Dear #NAME#," & vbCrlf & vbCrlf _
	& "Your order was updated to : #STAT#" & vbCRlf & vbCrlf _
	& "-----------------------------------------------------" & vbCrLf _
	& "Name : #NAME#" & vbCrLf _
	& "Order Number : #ORDER#" & vbCrLf _
	& "Order Date : #DATE#" & vbCrLf _
	& "Order Total : #TOTAL#" & vbCrLf _
	& "-----------------------------------------------------" & vbCRlf & vbCRlf _
	& "Thank you for your support." & vbCrlf & vbCrlf _
	& "#STORE#" & vbCrLf & vbCrLf
end if
if isNull(paySuccessMsg) or isEmpty(paySuccessMsg) then
	paySuccessMsg = "" _
	& "As soon as we have received and verified your payment, " _
	& "we will complete your order and update your order's " _
	& "status. Please note that you can check on your order's " _
	& "status by logging on to your Account and clicking on " _
	& "the order.<br><br>" & vbCrLf & vbCrLf _
	& "Thank you for your support.<br><br>" & vbCrLf & vbCrLf _
	& "<b>#STORE#</b>"
end if
if isNull(payErrorMsg) or isEmpty(payErrorMsg) then
	payErrorMsg = "" _
	& "It appears that there was an error while you were " _
	& "attempting to submit a payment for this order, or the " _
	& "payment was cancelled. If there was an error " _
	& "while trying to pay for this Order, you can Log On to " _
	& "your Account where you will be able to re-attempt payment " _
	& "for this Order.<br><br>" & vbCrLf & vbCrLf _
	& "If you cancelled due to a problem or if you were not happy " _
	& "with something, please let us know at #SALES#<br><br>" & vbCrLf & vbCrLf _
	& "Thank you for your support.<br><br>" & vbCrLf & vbCrLf _
	& "<b>#STORE#</b>"
end if
if isNull(passRequestEmail) or isEmpty(passRequestEmail) then
	passRequestEmail = "" _
	& "Dear #NAME#," & vbCRlf & vbCRlf _
	& "Here is the result of your query :" & vbCRlf & vbCRlf _
	& "------------------------------------------" & vbCrLf _
	& "Password : #PASS#" & vbCrLf _
	& "------------------------------------------" & vbCrLf & vbCRlf _
	& "Thank you for your interest in our store." & vbCrLf & vbCRlf _
	& "Regards" & vbCrLf & vbCrLf _
	& "#STORE#"
end if
if isNull(emailToFriend) or isEmpty(emailToFriend) then
	emailToFriend = "" _
	& "Hi, #STORE# has this item that I thought you would " _
	& "really like to know about." & vbCrLf & vbCrLf _
	& "#PRICE# - #PROD#" & vbCrLf & vbCrLf _
	& "Click on the link to see more info." & vbCrLf & vbCrLf _
	& "#LINK#" & vbCrLf
end if
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Text Configuration</font></b>
</P>

<form method="post" action="utilities_textexec.asp" name="TextMod">

<!-- termsAndCond -->
<span class="textBlockHead">Terms and Conditions</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" valign=top>
		Enter the general Terms and Conditions for the store here. You 
		may use HTML tags (see example). Be sure to specify things such 
		as your store's payment policy, return policy, and any other 
		rules and conditions you may want to add.
	</td>
	<td bgcolor="#EEEEEE" valign=top>
		<textarea name="termsAndCond" cols="55" rows="20"><%=server.HTMLEncode(termsAndCond)%></textarea>
	</td>
</tr>
</table>

<br>

<!-- saveOrderEmail -->
<span class="textBlockHead">Submit Order Email Body</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" valign=top>
		Enter the Email message that you want to send to your customers 
		when they submit the order to you. Use only plain text, no 
		HTML (some EMail software may not display HTML correctly).
		<br><br>
		Below is a list of "tags" that can be used to personalize the 
		email. These tags will be replaced with their actual values 
		before the email is sent out :
		<br><br>
		<code>#NAME#&nbsp;</code> - Customer name<br>
		<code>#DATE#&nbsp;</code> - Current date<br>
		<code>#ORDER#</code> - Order number<br>
		<code>#TOTAL#</code> - Order total<br>
		<code>#ITEMS#</code> - Order item detail<br>
		<code>#COMMT#</code> - Customer comments<br>
		<code>#PAYMT#</code> - Payment Method<br>
		<code>#STORE#</code> - Store name<br>
		<code>#SALES#</code> - Sales Email addr<br>
	</td>
	<td bgcolor="#EEEEEE" valign=top>
		<textarea name="saveOrderEmail" cols="55" rows="20"><%=server.HTMLEncode(saveOrderEmail)%></textarea>
	</td>
</tr>
</table>

<br>

<!-- statusUpdateEmail -->
<span class="textBlockHead">Update Order Status Email Body</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" valign=top>
		Enter the Email message that you want to send to your customers 
		when the order status is updated. Use only plain text, no 
		HTML (some EMail software may not display HTML correctly).
		<br><br>
		Below is a list of "tags" that can be used to personalize the 
		email. These tags will be replaced with their actual values 
		before the email is sent out :
		<br><br>
		<code>#NAME#&nbsp;</code> - Customer name<br>
		<code>#STAT#&nbsp;</code> - Order status<br>
		<code>#ORDER#</code> - Order number<br>
		<code>#DATE#&nbsp;</code> - Order date<br>
		<code>#TOTAL#</code> - Order total<br>
		<code>#STORE#</code> - Store name<br>
		<code>#SALES#</code> - Sales Email addr<br>
	</td>
	<td bgcolor="#EEEEEE" valign=top>
		<textarea name="statusUpdateEmail" cols="55" rows="20"><%=server.HTMLEncode(statusUpdateEmail)%></textarea>
	</td>
</tr>
</table>

<br>

<!-- passRequestEmail -->
<span class="textBlockHead">Password Request Email Body</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" valign=top>
		Enter the Email message that you want to send to your customers 
		when they request their password in your store. Use only plain 
		text, no HTML (some EMail software may not display HTML 
		correctly).
		<br><br>
		Below is a list of "tags" that can be used to personalize the 
		email. These tags will be replaced with their actual values 
		before the email is sent out :
		<br><br>
		<code>#NAME#&nbsp;</code> - Customer name<br>
		<code>#PASS#&nbsp;</code> - Password<br>
		<code>#STORE#</code> - Store name<br>
	</td>
	<td bgcolor="#EEEEEE" valign=top>
		<textarea name="passRequestEmail" cols="55" rows="20"><%=server.HTMLEncode(passRequestEmail)%></textarea>
	</td>
</tr>
</table>

<br>

<!-- emailToFriend -->
<span class="textBlockHead">Email To Friend Email Body</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" valign=top>
		Enter the default Email message that you want to be sent when 
		a customer uses the "Email To a Friend" function. Use 
		only plain text, no HTML (some EMail software may not display 
		HTML correctly).
		<br><br>
		Below is a list of "tags" that can be used to customize the 
		email. These tags will be replaced with their actual values 
		before the email is sent out :
		<br><br>
		<code>#PROD#&nbsp;</code> - Product Description<br>
		<code>#LINK#&nbsp;</code> - Link to Product<br>
		<code>#PRICE#</code> - Product Price<br>
		<code>#STORE#</code> - Store name<br>
	</td>
	<td bgcolor="#EEEEEE" valign=top>
		<textarea name="emailToFriend" cols="55" rows="20"><%=server.HTMLEncode(emailToFriend)%></textarea>
	</td>
</tr>
</table>

<br>

<!-- paySuccessMsg -->
<span class="textBlockHead">Successful Payment Message</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" valign=top>
		This message is displayed to the customer upon return from 
		a payment processor, if payment was determined to be successful. 
		You may use HTML tags (see example) to further enhance the 
		message's appearance.
		<br><br>
		Below is a list of replacement "tags" that can be used. These 
		tags will be replaced with their actual values before the 
		message is displayed :
		<br><br>
		<code>#STORE#</code> - Store name<br>
		<code>#SALES#</code> - Sales Email addr<br>
	</td>
	<td bgcolor="#EEEEEE" valign=top>
		<textarea name="paySuccessMsg" cols="55" rows="20"><%=server.HTMLEncode(paySuccessMsg)%></textarea>
	</td>
</tr>
</table>

<br>

<!-- payErrorMsg -->
<span class="textBlockHead">Unsuccessful Payment Message</span><br>
<table border=0 cellspacing=0 cellpadding=3 width="100%" class="textBlock">
<tr>
	<td bgcolor="#EEEEEE" valign=top>
		This message is displayed to the customer upon return from 
		a payment processor, if payment was determined to be unsuccessful. 
		You may use HTML tags (see example) to further enhance the 
		message's appearance.
		<br><br>
		Below is a list of replacement "tags" that can be used. These 
		tags will be replaced with their actual values before the 
		message is displayed :
		<br><br>
		<code>#STORE#</code> - Store name<br>
		<code>#SALES#</code> - Sales Email addr<br>
	</td>
	<td bgcolor="#EEEEEE" valign=top>
		<textarea name="payErrorMsg" cols="55" rows="20"><%=server.HTMLEncode(payErrorMsg)%></textarea>
	</td>
</tr>
</table>

<br>

<center>
	<input type="submit" name="submit1" value="Update Configuration">
</center>

</form>

<!--#include file="_INCfooter_.asp"-->
