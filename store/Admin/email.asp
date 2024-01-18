<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : send Email
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
const adminLevel = 1
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<%
'Email
dim emailFrom
dim emailTo
dim emailToName
dim emailSubj
dim emailBody

'Database
dim mySQL, cn, rs

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Close Database Connection
call closedb()

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Send Email</font></b>
	<br><br>
</P>

<%
'Pre-Populate some fields
emailTo     = trim(request.QueryString("emailTo"))
emailToName = trim(request.QueryString("emailToName"))
emailSubj   = trim(request.QueryString("emailSubj"))
emailBody   = trim(request.QueryString("emailBody"))
if len(emailBody) = 0 then
	emailBody = "" _
		& "TO : " & emailToName & vbCrLf &vbCrLf _
		& "RE : " & emailSubj   & vbCrLf
end if

'Are we returning to this page with a message?
if len(trim(Request.QueryString("msg"))) > 0 then
%>
	<font color=red><%=Request.QueryString("msg")%></font>
	<br><br>
<%
end if

'Send email via user's email client
if mailComp = "0" then
%>
<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<TR>
		<TD>
			The store is currently configured NOT to send email via 
			an email component from the web server. To enable web 
			based email for this store, configure your email server 
			and component in your store's configuration settings.
<%
			if len(emailTo) > 0 then
%>			
			<br><br>
			You can also send an email to the address below with your 
			regular Email Client software (eg. Outlook Express, Eudora, 
			etc.).
			
			<br><br>
			Click Email Address : <a href="mailto:<%=emailTo%>?subject=<%=emailSubj%>"><%=emailTo%></a>
<%
			end if
%>
			<br><br>
		</TD>
	</TR>
</TABLE>
<%

'Send email via an email component
else
%>
<span class="textBlockHead">Enter Email info and click 'Send'...</span><br>
<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
<form METHOD="POST" name="sendEmail" action="email_exec.asp">
	<TR>
		<TD nowrap><b>From Email</b> </TD>
		<TD>
			<select name=emailFrom id=emailFrom size=1>
				<option value="">---- Select ----</option>
				<option value="<%=pEmailSales%>"><%=pEmailSales%></option>
				<option value="<%=pEmailAdmin%>"><%=pEmailAdmin%></option>
			</select>
		</TD>
	</TR>
	<TR>
		<TD nowrap><b>To Email</b> </TD>
		<TD><input type="text" name="emailTo" size="40" maxlength="100" value="<%=emailTo%>"></TD>
	</TR>	
	<TR>
		<TD nowrap><b>Subject</b> </TD>
		<TD><input type="text" name="emailSubj" size="40" maxlength="200" value="<%=emailSubj%>"></TD>
	</TR>	
	<TR>
		<TD align="left" colspan="2" nowrap>
			<b>Message</b> <br>
			<textarea name=emailBody cols=55 rows=9><%=emailBody%></textarea>
		</TD>
	</TR>
	<TR>
		<TD align="center" colspan="2" nowrap>
			<input type="checkbox" name="contType" value="1"> 
			Send as HTML Email
		</TD>
	</TR>
	<TR>
		<TD colspan="2" align=center valign=middle>
			<input type="SUBMIT" name="Submit" value="Send Email">
			<br><br>
		</TD>
	</TR>
</form>
</TABLE>
<%
end if
%>
		
<!--#include file="_INCfooter_.asp"-->