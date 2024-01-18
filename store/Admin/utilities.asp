<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Administration Home Page
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
<%
'Work Fields
dim mySQL

'*************************************************************************

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Setup & Utilities</font></b>
	<br><br>
</P>

<span class="textBlockHead">Setup & Configuration</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr>
	<td>
			
		<a href="utilities_config.asp">Store Configuration</a> - 
		<i>Modify your store's general configuration settings.</i><br><br>
				
		<a href="utilities_text.asp">Text Configuration</a> - 
		<i>Modify some of the text used in your store.</i><br><br>

	</td>
</tr>
</table>

<br>

<span class="textBlockHead">Test and Repair</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr>
	<td>
			
		<a href="utilities_DBwrite.asp">Test Database Read and Write</a> - 
		<i>Check if you can connect to the database and make modifications.</i><br><br>

		<a href="utilities_DBstruc.asp">Test Database Structure</a> - 
		<i>Check if database has required files and fields, and repair if necessary.</i><br><br>

		<a href="utilities_Email.asp">Test Email</a> - 
		<i>Check if you are able to send emails from your store.</i><br><br>

	</td>
</tr>
</table>

<br>
		
<span class="textBlockHead">General</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr>
	<td>
			
		<a href="upload.asp">Upload Files</a> - 
		<i>Upload Product Images and/or Downloadable Products to your web server.</i><br><br>

		<a href="email.asp">Send Email</a> - 
		<i>Send emails from your store.</i><br><br>

		<a href="utilities_ServerVars.asp">Display Server Variables</a> - 
		<i>Display your web server's variables.</i><br><br>
				
	</td>
</tr>
</table>

<br>

<span class="textBlockHead">Other Utilities</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr>
	<td>

		List  
		<a href="utilities_SQL.asp?mySQL=<%=Server.UrlEncode("SELECT * FROM products")  %>">Products</a>, 
		<a href="utilities_SQL.asp?mySQL=<%=Server.UrlEncode("SELECT * FROM categories")%>">Categories</a>, 
		<a href="utilities_SQL.asp?mySQL=<%=Server.UrlEncode("SELECT * FROM options")   %>">Options</a>, 
		<a href="utilities_SQL.asp?mySQL=<%=Server.UrlEncode("SELECT * FROM locations") %>">Locations</a>, 
		<a href="utilities_SQL.asp?mySQL=<%=Server.UrlEncode("SELECT * FROM shipRates") %>">ShipRates</a>, 
		<a href="utilities_SQL.asp?mySQL=<%=Server.UrlEncode("SELECT * FROM cartHead")  %>">CartHead</a>, 
		<a href="utilities_SQL.asp?mySQL=<%=Server.UrlEncode("SELECT * FROM customer")  %>">Customer</a>, 
		<a href="utilities_SQL.asp?mySQL=<%=Server.UrlEncode("SELECT * FROM DiscOrder") %>">Discounts</a> 
		<br><br>

		SQL Command - <i>Send raw SQL command 
		to the database.</i><br>
		<%
		mySql = trim(request.form("mySQL"))
		if len(mySQL) = 0 then
			mySql = request.querystring("mySQL")
		end if
		%>
		<form method="post" action="utilities_SQL.asp" name="form1">
			<textarea name="mySQL" cols="60" rows="5"><%=mySQL%></textarea>
			<br>
			<input type="submit" name="Submit" value="Execute">
		</form>

	</td>
</tr>
</table>

<!--#include file="_INCfooter_.asp"--> 
