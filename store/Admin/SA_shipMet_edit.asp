<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Shipping Method Maintenance
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
<!--#include file="_INCshipping_.asp"-->
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="../Scripts/_INCappFunctions_.asp"-->
<%
'Database
dim mySQL, cn, rs

'ShipMethod
dim idShipMethod
dim shipDesc
dim status

'Work Fields
dim action

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if action <> "edit" and action <> "add" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

if action = "edit" then

	'Get idShipMethod
	idShipMethod = trim(Request.QueryString("recId"))
	if len(idShipMethod) = 0 then
		idShipMethod = trim(Request.Form("recId"))
	end if
	if idShipMethod = "" or not isNumeric(idShipMethod) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if

	'Get ShipMethod Record
	mySQL="SELECT shipDesc,status " _
	    & "FROM   ShipMethod " _
	    & "WHERE  idShipMethod = " & idShipMethod
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	else
		shipDesc = rs("shipDesc")
		status   = rs("status")
	end if
	call closeRS(rs)
	
end if

'Close database connection
call closedb()
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Shipping - Store Method Maintenance</font></b>
	<br><br>
</P>

<%
'Page Tabs
call shipTabs("CM")

'Edit
if action = "edit" then
	if len(trim(Request.QueryString("msg"))) > 0 then
%>
		<font color=red><%=Request.QueryString("msg")%></font>
		<br><br>
<%
	end if
%>
	<span class="textBlockHead">Edit Shipping Method</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td><b>Description</b></td>
			<td><b>Status</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_shipMet_exec.asp" name="form1">
			<td>
				<input type=text name=shipDesc id=shipDesc size=25 maxlength=100 value="<%=shipDesc%>">
			</td>
			<td>
				<select name=status id=status size=1>
					<option value="A" <%=checkMatch(status,"A")%>>Active</option>
					<option value="I" <%=checkMatch(status,"I")%>>InActive</option>
				</select>
			</td>
			<td>
				<input type=hidden name=idShipMethod id=idShipMethod value="<%=idShipMethod%>">
				<input type=hidden name=action       id=action       value="edit">
				<input type=submit name=submit1      id=submit1      value="Update">
			</td>
			</form>
		</tr>
	</table>
<%
end if

'Add
if action = "add" then
%>
	<span class="textBlockHead">Add Shipping Method</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td><b>Description</b></td>
			<td><b>Status</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_shipMet_exec.asp" name="form2">
			<td>
				<input type=text name=shipDesc id=shipDesc size=20 maxlength=100>
			</td>
			<td>
				<select name=status id=status size=1>
					<option value="A">Active</option>
					<option value="I">InActive</option>
				</select>
			</td>
			<td>
				<input type=hidden name=action  id=action  value="add">
				<input type=submit name=submit1 id=submit1 value="Add New Method">
			</td>
			</form>
		</tr>
	</table>
<%
end if

if action = "edit" or action = "add" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>
		
		<b>Description</b> - Mandatory. The description you want to give 
		to this Shipping Method. Eg. 'FedEx Ground', 'UPS 2 Day Air', 
		'MyStore Special Shipping', etc.<br><br>
		
		<b>Status</b> - Mandatory. If set to 'Yes' this Shipping Method 
		will be active and therefore included when the system calculates 
		Shipping Rates for a particular order.<br><br>
		
	</td></tr>
	</table>
<%
end if
%>
<!--#include file="_INCfooter_.asp"-->
<%
'*********************************************************************
'Create Navigation Links
'*********************************************************************
sub maintNavLinks()
%>
	[ 
	<a href=SA_shipMet.asp>List Methods</a> | 
	<a href=SA_shipMet_edit.asp?action=add>Add</a> 
	]
<%
end sub
%>
