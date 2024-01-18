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
'*************************************************************************

'Are we in demo mode?
if demoMode = "Y" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("DEMO MODE. Sorry, this featured is NOT available in Demo Mode.")
end if

'Get variables from the Form
termsAndCond      = trim(Request.Form("termsAndCond"))
saveOrderEmail    = trim(Request.Form("saveOrderEmail"))
paySuccessMsg     = trim(Request.Form("paySuccessMsg"))
payErrorMsg       = trim(Request.Form("payErrorMsg"))
passRequestEmail  = trim(Request.Form("passRequestEmail"))
emailToFriend     = trim(Request.Form("emailToFriend"))
statusUpdateEmail = trim(Request.Form("statusUpdateEmail"))

'termsAndCond
if len(termsAndCond) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Terms and Conditions can not be empty.")
end if

'saveOrderEmail
if len(saveOrderEmail) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Submit Order Email Body can not be empty.")
end if

'statusUpdateEmail
if len(statusUpdateEmail) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Update Order Status Email Body can not be empty.")
end if

'passRequestEmail
if len(passRequestEmail) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Password Request Email Body can not be empty.")
end if

'emailToFriend
if len(emailToFriend) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Email To Friend Email Body can not be empty.")
end if

'paySuccessMsg
if len(paySuccessMsg) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Successful Payment Message can not be empty.")
end if

'payErrorMsg
if len(payErrorMsg) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Unsuccessful Payment Message can not be empty.")
end if

'Update the text configuration file
call writeConfigFile()

%>
<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Text Configuration</font></b>
</P>

<table border=0 cellspacing=0 cellpadding=10 width="100%" class="textBlock">
<tr><td>

	<br>

	<b><font size=3 color=green>Success!</font></b>

	<br><br><br>

	<font size=2>
		Text Configurations were successfully updated.
	</font>

	<br><br>
	
</td></tr>
</table>

<!--#include file="_INCfooter_.asp"-->
<%
'**********************************************************************
'Update Text Configurations
'**********************************************************************
sub writeConfigFile()

	'Create an array of variables which will use to drive the update
	dim configArr, I
	configArr = "termsAndCond,saveOrderEmail,paySuccessMsg," _
		& "payErrorMsg,passRequestEmail,emailToFriend," _
		& "statusUpdateEmail"
	configArr = split(configArr,",")
	
	'Open Database
	call openDb()

	'Loop through array and UPDATE / INSERT new config settings
	for I = 0 to UBound(configArr)
		if len(trim(configArr(I))) > 0 then
			mySQL = "SELECT configVar " _
				&   "FROM   storeAdmin " _
				&   "WHERE  configVar='" & configArr(I) & "' " _
				&   "AND    adminType='T'"
			set rs = openRSexecute(mySQL)
			if rs.EOF then
				'INSERT
				call closeRS(rs)
				mySQL = "INSERT INTO storeAdmin " _
					  & "(adminType,configVar,configValLong) " _
					  & "VALUES " _
					  & "('T','" & configArr(I) & "','" & replace(eval(configArr(I)),"'","''") & "')"
				set rs = openRSexecute(mySQL)
			else
				'UPDATE
				call closeRS(rs)
				mySQL = "UPDATE storeAdmin SET " _
				      & "       configValLong='" & replace(eval(configArr(I)),"'","''") & "' " _
					  & "WHERE  configVar='" & configArr(I) & "' " _
					  & "AND    adminType='T'"
				set rs = openRSexecute(mySQL)
			end if
		end if
	next

	'Close Database
	call closedb()
	
end sub

%>