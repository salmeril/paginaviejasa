<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Upload Files
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
'Database
dim mySQL, cn, rs

'*************************************************************************

'Are we in demo mode?
if demoMode = "Y" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("DEMO MODE. Sorry, this featured is NOT available in Demo Mode.")
end if

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Close Database Connection
call closedb()

'Get UserID of locally logged on user (if applicable)
dim logonWinUser
logonWinUser = Request.ServerVariables("LOGON_USER")
%>

<!--#include file="_INCheader_.asp"-->

<b><font size=3>Upload Files</font></b><br>
<img src="x_cleardot.gif" border=0 height=5><br>
<%
if len(trim(Request.QueryString("msg"))) > 0 then
	Response.Write "<font color=red>" & Request.QueryString("msg") & "</font><br>"
else
	Response.Write "&nbsp;<br>"
end if
%>
<img src="x_cleardot.gif" border=0 height=5><br>

<span class="textBlockHead">1. Select Files to Upload.</span><br>
<table border=0 cellspacing=0 cellpadding=5 width=350 class="textBlock">
<form method="post" action="upload_exec.asp" enctype="multipart/form-data" id="uplform" name="uplform">
	<tr>
		<td nowrap>File 01</td>
		<td><input name="file01" size="30" type="file"></td>
	</tr>
	<tr>
		<td nowrap>File 02</td>
		<td><input name="file02" size="30" type="file"></td>
	</tr>
	<tr>
		<td nowrap>File 03</td>
		<td><input name="file03" size="30" type="file"></td>
	</tr>
	<tr>
		<td nowrap>File 04</td>
		<td><input name="file04" size="30" type="file"></td>
	</tr>
	<tr>
		<td nowrap>File 05</td>
		<td><input name="file05" size="30" type="file"></td>
	</tr>
</table>

<br>

<span class="textBlockHead">2. Select Upload Folder.</span><br>
<table border=0 cellspacing=0 cellpadding=5 width=350 class="textBlock">
	<tr>
		<td>
			<input type="radio" name="uplFolder" value="<%=pImagesDir%>" checked>Product Images Directory (<b><%=pImagesDir%></b>)<br>
			<input type="radio" name="uplFolder" value="<%=pDownloadDir%>">Downloadable Items Directory (<b><%=pDownloadDir%></b>)<br>
		</td>
	</tr>
</table>

<br>

<span class="textBlockHead">3. Submit Upload.</span><br>
<table border=0 cellspacing=0 cellpadding=5 width=350 class="textBlock">
	<tr>
		<td align=center valign=middle>
			<br>
			<input type="submit" name="submit1" value="Upload Files"><br><br>
<%
			'See if user is already logged on locally to the web server
			If IsEmpty(logonWinUser) Or IsNull(logonWinUser) Or logonWinUser="" Then
%>
				<input type="checkbox" name="logonWin" value="Y">Logon locally to web server (<a href="#logonWin">Help</a>)
<%
			else
%>
				<b>You are logged on to your server as <font color=red><%=logonWinUser%></font></b>
<%
			end if
%>
			<br><br>
		</td>
	</tr>
</form>
</table>

<br>

<span class="textBlockHead">Help and Instructions</span><br>
<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
<tr><td>
	<a name="logonWin"></a>
	1. You can upload up to 5 files at a time to the folder of your 
	choice.<br><br>

	2. For the utility to save the files on your web server, you must 
	have the correct permissions. The selected folder must either have 
	<b>read/write</b> permissions given to the anonymous web user account, 
	OR you must 'Log on Locally' to your web server with another 
	UserID and Password that has these permissions. Users who have 
	direct access to their web servers can set this up themselves, 
	or you will have to contact your web hosting company to do this 
	for you. NOTE : Most users with hosted web accounts will probably 
	already have an account that they can use to upload files to the 
	server. This account is usually supplied by the hosting company 
	to the user for the purpose of uploading web pages and making 
	modifications to their web site. This account will usually have the 
	permissions required by the web server to allow uploads.<br><br>

	3. In addition to using this utility, you can FTP the files to 
	their appropriate folders on your web server.<br><br>

</td></tr>
</table>

<!--#include file="_INCfooter_.asp"-->
