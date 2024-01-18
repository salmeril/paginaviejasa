<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Product Review Maintenance
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
<%
'Database
dim mySQL, cn, rs

'Reviews
dim idReview
dim idProduct
dim revDate
dim revDateInt
dim revAuditInfo
dim revStatus
dim revRating
dim revName
dim revLocation
dim revEmail
dim revSubj
dim revDetail

'Products
dim description

'Work Fields
dim I
dim action

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Review Maintenance</font></b>
	<br><br>
</P>

<%
'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if action <> "edit" and action <> "del" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idReview
idReview = trim(Request.QueryString("recId"))
if len(idReview) = 0 then
	idReview = trim(Request.Form("recId"))
end if
if idReview = "" or not isNumeric(idReview) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Review ID.")
end if

'Get Review details
mySQL="SELECT idProduct,revDate,revAuditInfo,revStatus,revRating," _
	& "       revName,revLocation,revEmail,revSubj,revDetail " _
    & "FROM   reviews " _
    & "WHERE  idReview = " & idReview
set rs = openRSexecute(mySQL)
if rs.eof then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Review ID.")
else
	revDetail	= rs("revDetail")
	idProduct	= rs("idProduct")
	revDate		= rs("revDate")
	revAuditInfo= rs("revAuditInfo")
	revStatus	= rs("revStatus")
	revRating	= rs("revRating")
	revName		= rs("revName")
	revLocation	= rs("revLocation")
	revEmail	= rs("revEmail")
	revSubj		= rs("revSubj")
end if
call closeRS(rs)
	
'Get Product info
mySQL="SELECT description " _
    & "FROM   products " _
    & "WHERE  idProduct = " & idProduct
set rs = openRSexecute(mySQL)
if rs.eof then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Product ID.")
else
	description	= rs("description")
end if
call closeRS(rs)

'Edit form
if action = "edit" then
	if len(trim(Request.QueryString("msg"))) > 0 then
%>
		<font color=red><%=Request.QueryString("msg")%></font>
		<br><br>
<%
	end if
%>
	<span class="textBlockHead">Edit Review</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<form method="post" action="SA_rev_exec.asp" name="form1">
		<tr>
			<td align=right nowrap><b>Date</b></td>
			<td align=left><%=formatTheDate(revDate)%></td>
		</tr>
		<tr>
			<td align=right nowrap><b>IP Address</b></td>
			<td align=left><a href="http://www.samspade.org/t/lookat?a=<%=revAuditInfo%>" target="_blank"><%=revAuditInfo%></a></td>
		</tr>
		<tr>
			<td align=right nowrap><b>Product</b></td>
			<td align=left><%=description%></td>
		</tr>
		<tr>
			<td align=right nowrap><b>Rating</b></td>
			<td align=left nowrap>
				<select name=revRating id=revRating size=1>
					<option value="1" <%=checkMatch(revRating,"1")%>>1</option>
					<option value="2" <%=checkMatch(revRating,"2")%>>2</option>
					<option value="3" <%=checkMatch(revRating,"3")%>>3</option>
					<option value="4" <%=checkMatch(revRating,"4")%>>4</option>
					<option value="5" <%=checkMatch(revRating,"5")%>>5</option>
				</select>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<b>Status</b>
				&nbsp;
				<select name=revStatus id=revStatus size=1>
					<option value="A" <%=checkMatch(revStatus,"A")%>>Active</option>
					<option value="I" <%=checkMatch(revStatus,"I")%>>Pending</option>
					<option value="R" <%=checkMatch(revStatus,"R")%>>Rejected</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Name</b></td>
			<td align=left>
				<input type=text name=revName id=revName size=30 maxlength=250 value="<%=server.HTMLEncode(revName & "")%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Location</b></td>
			<td align=left>
				<input type=text name=revLocation id=revLocation size=30 maxlength=250 value="<%=server.HTMLEncode(revLocation & "")%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Email</b></td>
			<td align=left>
				<input type=text name=revEmail id=revEmail size=30 maxlength=100 value="<%=server.HTMLEncode(revEmail & "")%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Subject</b></td>
			<td align=left>
				<input type=text name=revSubj id=revSubj size=30 maxlength=100 value="<%=server.HTMLEncode(revSubj & "")%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Review</b></td>
			<td align=left>
				<textarea name="revDetail" rows=10 cols="50" wrap="soft"><%=server.HTMLEncode(revDetail & "")%></textarea>
			</td>
		</tr>
		<tr>
			<td colspan=2 align=center valign=middle>
				<br>
				<input type=hidden name=idReview id=idReview value="<%=idReview%>">
				<input type=hidden name=action   id=action   value="edit">
				<input type=submit name=submit1  id=submit1  value="Update Review">
				<br><br>
			</td>
		</tr>
		</form>
	</table>
<%
end if

'Delete review
if action = "del" then
%>
	<span class="textBlockHead">Delete Review</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<form method="post" action="SA_rev_exec.asp" name="form1">
		<tr>
			<td align=right nowrap><b>Date</b></td>
			<td align=left><%=formatTheDate(revDate)%></td>
		</tr>
		<tr>
			<td align=right valign=top nowrap><b>Product</b></td>
			<td align=left><%=description%></td>
		</tr>
		<tr>
			<td align=right valign=top nowrap><b>Name</b></td>
			<td align=left><%=server.HTMLEncode(revName & "")%></td>
		</tr>
		<tr>
			<td align=right valign=top nowrap><b>Subject</b></td>
			<td align=left><%=server.HTMLEncode(revSubj & "")%></td>
		</tr>
		<tr>
			<td align=right valign=top nowrap><b>Review</b></td>
			<td align=left><%=server.HTMLEncode(revDetail & "")%></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td valign=middle>
				<br>
				<input type=hidden name=idReview id=idReview value="<%=idReview%>">
				<input type=hidden name=action   id=action   value="del">
				<input type=submit name=submit1  id=submit1  value="Delete Review">
				<br><br>
			</td>
		</tr>
		</form>
	</table>
<%
end if

if action = "edit" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>
	
		<b>Note : All fields are mandatory.</b><br><br>
	
		<b>Rating</b> - 1 = Worst, 5 = Best<br><br>
	
		<b>Status</b> - "Active" reviews are viewable by your customers, 
		and will therefore also be included when calculating the average 
		rating for the product. "Pending" reviews are awaiting authorization, 
		and are therefore not yet viewable. "Rejected" reviews have been 
		deliberately made inactive, and will therefore not be viewable. 
		Rejected reviews are included when the system checks for duplicate 
		IP addresses.<br><br>
	
		<b>Name</b> - Reviewer's full name.<br><br>

		<b>Location</b> - City, State or Country of reviewer.<br><br>
	
		<b>Email</b> - Email address of reviewer.<br><br>
	
		<b>Subject</b> - Short and descriptive heading for the 
		review.<br><br>
	
		<b>Review</b> - Full text of the review.<br><br>
	
	</td></tr>
	</table>
<%
end if

'Close database
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
	<a href=SA_rev.asp?recallCookie=1>List Reviews</a> | 
	<a href=SA_rev_edit.asp?action=edit&recid=<%=idReview%>>Edit</a> | 
	<a href=SA_rev_edit.asp?action=del&recid=<%=idReview%>>Delete</a> 
	]
<%
end sub
%>