<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Email a Product To a Friend
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
<!--#include file="_INCappEmail_.asp"-->
<%
'Product
dim idProduct
dim description
dim price

'Email
dim emailName
dim emailTo
dim emailBody

'Work Fields
dim arrayErrors
dim formID

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

'Are we calling this page from the products page?
if Request.Form("formID") = "" then

	'Get Product Details
	idProduct   = validHTML(Request.QueryString("idProduct"))
	description = validHTML(Request.QueryString("description"))
	price       = validHTML(Request.QueryString("price"))
	
	'Build Email Body
	emailBody = ""
	mySQL = "SELECT configValLong " _
		&   "FROM   storeAdmin " _
		&   "WHERE  configVar = 'emailToFriend' " _
		&   "AND    adminType = 'T'"
	set rsTemp = openRSexecute(mySQL)
	if not rstemp.eof then
		emailBody = trim(rsTemp("configValLong"))
	end if
	call closeRS(rsTemp)
	
	'Check for tags and replace
	emailBody = replace(emailBody,"#PROD#",description)
	emailBody = replace(emailBody,"#LINK#",urlNonSSL & "prodView.asp?idProduct=" & idProduct)
	emailBody = replace(emailBody,"#PRICE#",pCurrencySign & moneyS(price))
	emailBody = replace(emailBody,"#STORE#",pCompany)

'This page has called itself
else

	'Get Form Fields
	idProduct = validHTML(request.Form("idProduct"))
	emailName = validHTML(request.Form("emailName"))
	emailTo   = validHTML(request.Form("emailTo"))
	emailBody = validHTML(request.Form("emailBody"))
	
	'Do some checks
	if len(emailName) = 0 then
		arrayErrors = arrayErrors & "|emailName"
	end if
	if len(emailTo) = 0 or invalidChar(emailTo,1,"@.-_") then
		arrayErrors = arrayErrors & "|emailTo"
	end if
	if len(emailBody) = 0 then
		arrayErrors = arrayErrors & "|emailBody"
	end if
	
	'If there was no errors, send the email.
	if len(trim(arrayErrors)) = 0 then
	
		'Send Email
		call sendmail (emailName, pEmailSales, emailTo, pCompany, emailBody, 0)

		'Say Thank You
		response.redirect "sysMsg.asp?msg=" & server.URLEncode(langGenEmailFriendMsg)
		
	end if
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<table border="0" cellspacing="0" cellpadding="2" width="100%">
	<tr>
		<td valign=middle class="CPpageHead">
			<b><%=langGenEmailFriendHdr%></b><br>
		</td>
	</tr>
	<tr>
		<form METHOD="POST" name="emailToFriend" action="emailToFriend.asp">
		<td valign=top>
			<br>
<%
			'If there were errors, show message
			if len(trim(arrayErrors)) > 0 then
				arrayErrors = split(LCase(arrayErrors),"|")
				Response.Write "<font color=red><i>" & langErrInvForm & "</i></font><br><br>"
			else
				arrayErrors = array("")
			end if
%>
			<%=langGenYourName & " " & checkFieldError("emailName",arrayErrors)%><br>
			<input type="text" name="emailName" size="20" maxlength="50" value="<%=emailName%>"><br>
			<%=langGenFriendEmail & " " & checkFieldError("emailTo",arrayErrors)%><br>
			<input type="text" name="emailTo" size="20" maxlength="50" value="<%=emailTo%>"><br>
			<%=langGenMessage & " " & checkFieldError("emailBody",arrayErrors)%><br>
<%
			'Check if customer is allowed to modify email body
			if pEmailFriendSec = -1 then
%>
				<textarea name="emailBodyDummy" rows=6 cols="40" readonly wrap="soft"><%=emailBody%></textarea><br><br>
				<input type="hidden" name="emailBody" value="<%=emailBody%>">				
<%
			else
%>
				<textarea name="emailBody" rows=6 cols="40" wrap="soft"><%=emailBody%></textarea><br><br>
<%
			end if
%>
			<input type="hidden" name="idProduct" value="<%=idProduct%>">
			<input type="hidden" name="formID"    value="00">
			<input type="SUBMIT" name="Submit"    value="<%=langGenSend%>">
		</td>
		</form>
	</tr>
</table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%
call closedb()
%>
