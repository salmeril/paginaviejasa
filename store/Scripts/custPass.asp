<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Retrieve a customers Password and send via email
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
<%
'Customer
dim status
dim Name
dim LastName
dim Email
dim Password

'Work Fields
dim arrayErrors
dim customerEmail
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

'Get Form ID
formID = Request.Form("formID")
if formID <> "00b" then
	formID = ""
end if

'Check if user clicked the "Get Password" button
if formID = "00b" then

	'Get email from the form
	Email = validHTML(request.form("Email"))

	'Do some checks
	if len(email) = 0 or invalidChar(Email,1,"@.-_") then
		arrayErrors = arrayErrors & "|email"
	else
		'Check if email exists on DB, and if Customer is still Active
		mySQL = "SELECT name, lastname, password " _
			  & "FROM   customer " _
			  & "WHERE  email = '" & validSQL(email,"A") & "' " _
			  & "AND    status = 'A'"
		set rsTemp = openRSexecute(mySQL)
		if not rstemp.eof then
		
			'Build Email Body
			customerEmail = ""
			mySQL = "SELECT configValLong " _
				&   "FROM   storeAdmin " _
				&   "WHERE  configVar = 'passRequestEmail' " _
				&   "AND    adminType = 'T'"
			set rsTemp2 = openRSexecute(mySQL)
			if not rstemp2.eof then
				customerEmail = trim(rsTemp2("configValLong"))
			end if
			call closeRS(rsTemp2)
	
			'Check for tags and replace
			customerEmail = replace(customerEmail,"#NAME#",rsTemp("name") & " " & rsTemp("lastName"))
			customerEmail = replace(customerEmail,"#PASS#",EnDeCrypt(Hex2Ascii(rsTemp("password")),rc4Key))
			customerEmail = replace(customerEmail,"#STORE#",pCompany)

			'Send Email to Customer
			call sendmail (pCompany, pEmailAdmin, email, langGenPassRequest, customerEmail, 0)
			
		else
		
			'Invalid email entered
			arrayErrors = arrayErrors & "|email"
			
		end if
		call closeRS(rsTemp)
	end if

	'There were no errors
	if len(trim(arrayErrors)) = 0 then
		response.redirect "sysMsg.asp?msg=" & server.URLEncode(langGenPassSentMsg)
	end if
end if
%>

<!--#include file="../UserMods/_INCtop_.asp"-->

<table border="0" cellspacing="0" cellpadding="2" width="350">
	<tr>
		<td valign=middle class="CPpageHead">
			<b><%=langGenPassRequest%></b><br>
		</td>
	</tr>
	<tr>
		<form METHOD="POST" name="custPass" action="custPass.asp">
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
			<%=langGenEmail & " " & checkFieldError("email",arrayErrors)%><br>
			<input type="text" name="email" size="20" maxlength="50" value="<%=email%>"><br><br>
			<input type="hidden" name="formID" value="00b">
			<input type="SUBMIT" name="Submit" value="<%=langGenSend%>">
		</td>
		</form>
	</tr>
</table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%
call closedb()
%>
