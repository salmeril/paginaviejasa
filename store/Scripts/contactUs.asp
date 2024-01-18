<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Contact Us page
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
'Email
dim emailTo
dim emailName
dim emailFrom
dim emailSubject
dim emailBody

'Work Fields
dim arrayErrors
dim formID

'Database
dim mySQL
dim conntemp
dim rstemp

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

'Get Form or QueryString parms
emailName	 = validHTML(Request("emailName"))
emailFrom	 = validHTML(Request("emailFrom"))
emailSubject = validHTML(Request("emailSubject"))
emailBody	 = validHTML(Request("emailBody"))

'Set To Email address
emailTo = pEmailSales

'Did the user click the "Send" button?
if Request.Form("formID") = "00" then

	'Do some checks
	if len(emailName) = 0 then
		arrayErrors = arrayErrors & "|emailName"
	end if
	
	if len(emailFrom) = 0 then
		arrayErrors = arrayErrors & "|emailFrom"
	else
		if inStr(emailFrom,"@") = 0 or inStr(emailFrom,".") = 0 then
			arrayErrors = arrayErrors & "|emailFrom"
		end if
		if invalidChar(emailFrom,1,"@.-_") then
			arrayErrors = arrayErrors & "|emailFrom"
		end if
	end if
	
	if len(emailSubject) = 0 then
		arrayErrors = arrayErrors & "|emailSubject"
	end if
	
	if len(emailBody) = 0 then
		arrayErrors = arrayErrors & "|emailBody"
	end if
	
	'If there was no errors, send the email.
	if len(trim(arrayErrors)) = 0 then
	
		'Send Email
		call sendmail (emailName, emailFrom, emailTo, emailSubject, emailBody, 0)

		'Say Thank You
		response.redirect "sysMsg.asp?msg=" & server.URLEncode(langGenContactUsMsg)
		
	end if
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<table border="0" cellspacing="0" cellpadding="2" width="100%">
	<tr>
		<td valign=middle class="CPpageHead">
			<b><%=langGenContactUsHdr%></b><br>
		</td>
	</tr>
	<tr>
		<form METHOD="POST" name="contactUs" action="contactUs.asp">
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
			<%=langGenTo%> : <b><%=emailTo%></b><br><br>

			<%=langGenYourName & " " & checkFieldError("emailName",arrayErrors)%><br>
			<input type="text" name="emailName" size="20" maxlength="50" value="<%=emailName%>"><br>

			<%=langGenEMail & " " & checkFieldError("emailFrom",arrayErrors)%><br>
			<input type="text" name="emailFrom" size="20" maxlength="50" value="<%=emailFrom%>"><br>
			
			<%=langGenSubject & " " & checkFieldError("emailSubject",arrayErrors)%><br>
			<input type="text" name="emailSubject" size="40" maxlength="50" value="<%=emailSubject%>"><br>

			<%=langGenMessage & " " & checkFieldError("emailBody",arrayErrors)%><br>

			<textarea name="emailBody" rows=6 cols="40" wrap="soft"><%=emailBody%></textarea><br><br>

			<input type="hidden" name="formID" value="00">
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
