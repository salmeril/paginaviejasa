<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Location Maintenance
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

'Locations
dim idLocation
dim locName
dim locCountry
dim locState
dim locTax
dim locShipZone
dim locStatus

'Work Fields
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
	<b><font size=3>Location & Tax Maintenance</font></b>
	<br><br>
</P>

<%
'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if  action <> "edit" _
and action <> "del" _
and action <> "add" _ 
and action <> "editstate" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idLocation
if action = "edit" or action = "del" or action = "editstate" then
	idLocation = trim(Request.QueryString("recId"))
	if len(idLocation) = 0 then
		idLocation = trim(Request.Form("recId"))
	end if
	if idLocation = "" or not isNumeric(idLocation) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if
end if

'Get Location Record
if action = "edit" or action = "del" or action = "editstate" then

	mySQL="SELECT locName,locCountry,locState,locTax," _
		& "       locShipZone,locStatus " _
	    & "FROM   Locations " _
	    & "WHERE  idLocation = " & idLocation
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	else
		locName		= rs("locName")
		locCountry	= rs("locCountry")
		locState	= rs("locState")
		locTax		= rs("locTax")
		locShipZone	= rs("locShipZone")
		locStatus   = rs("locStatus")
	end if
	call closeRS(rs)
	
end if

'Check for valid Country record
if action = "edit" or action = "del" then
	if not(IsNull(locState) or locState = "") then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Country Record ID.")
	end if
end if

'Check for valid State record
if action = "editstate" then
	if IsNull(locState) or locState = "" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid State/Province Record ID.")
	end if
end if

'Edit
if action = "edit" then
	if len(trim(Request.QueryString("msg"))) > 0 then
%>
		<font color=red><%=Request.QueryString("msg")%></font>
		<br><br>
<%
	end if
%>
	<span class="textBlockHead">Edit Country</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td nowrap><b>Code</b></td>
			<td nowrap><b>Country Name</b></td>
			<td nowrap><b>Tax %</b></td>
			<td nowrap><b>Zone</b></td>
			<td nowrap><b>Status</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_loc_exec.asp" name="editCountry">
			<td><input type=text name=locCountry  id=locCountry  size=2  maxlength=2   value="<%=locCountry%>"></td>
			<td><input type=text name=locName     id=locName     size=30 maxlength=100 value="<%=locName%>"></td>
			<td><input type=text name=locTax      id=locTax      size=5  maxlength=10  value="<%=formatNumber(locTax,2)%>"></td>
			<td><input type=text name=locShipZone id=locShipZone size=2  maxlength=2   value="<%=locShipZone%>"></td>
			<td>
				<select name=locStatus id=locStatus size=1>
					<option value="A" <%=checkMatch(locStatus,"A")%>>Active</option>
					<option value="I" <%=checkMatch(locStatus,"I")%>>InActive</option>
				</select>
			</td>
			<td>
				<input type=hidden name=idLocation id=idLocation value="<%=idLocation%>">
				<input type=hidden name=action     id=action     value="edit">
				<input type=submit name=submit1    id=submit1    value="Update">
			</td>
			</form>
		</tr>
		<tr><td colspan=6>&nbsp;</td></tr>
		<tr>
			<td colspan=6 bgcolor="#dddddd">
				<span class="textBlockHead">Add State or Province for <%=locName%></span>
			</td>
		</tr>
		<tr>
			<td nowrap><b>Code</b></td>
			<td nowrap><b>State/Province&nbsp;Name</b></td>
			<td nowrap><b>Tax&nbsp;%</b></td>
			<td nowrap><b>Zone</b></td>
			<td nowrap><b>Status</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_loc_exec.asp" name="addState">
			<td><input type=text name=locState    id=locState    size=2  maxlength=2></td>
			<td><input type=text name=locName     id=locName     size=30 maxlength=100></td>
			<td><input type=text name=locTax      id=locTax      size=5  maxlength=10></td>
			<td><input type=text name=locShipZone id=locShipZone size=2  maxlength=2></td>
			<td>
				<select name=locStatus id=locStatus size=1>
					<option value="A">Active</option>
					<option value="I">InActive</option>
				</select>
			</td>
			<td>
				<input type=hidden name=locCountry id=locCountry value="<%=locCountry%>">
				<input type=hidden name=action     id=action     value="addState">
				<input type=submit name=submit1    id=submit1    value=" Add ">
			</td>
			</form>
		</tr>
		<tr><td colspan=6>&nbsp;</td></tr>
		<tr>
			<td colspan=6 bgcolor="#dddddd">
				<span class="textBlockHead">Edit State or Province for <%=locName%></span>
			</td>
		</tr>
<%
		mySQL="SELECT idLocation,locName,locCountry,locState," _
			& "       locTax,locShipZone,locStatus " _
		    & "FROM   Locations " _
		    & "WHERE  locCountry = '" & locCountry & "' " _
		    & "AND    NOT(locState IS NULL OR locState = '') " _
		    & "ORDER BY locName "
		set rs = openRSexecute(mySQL)
		if rs.EOF then
%>
			<tr>
				<td align=center valign=middle colspan=6>
					<b>No States or Provinces.</b>
				</td>
			</tr>
<%
		else
%>
			<tr>
				<td nowrap><b>Code</b></td>
				<td nowrap><b>State/Province Name</b></td>
				<td nowrap><b>Tax %</b></td>
				<td nowrap><b>Zone</b></td>
				<td nowrap><b>Status</b></td>
				<td>&nbsp;</td>
			</tr>
<%
		end if
		do while not rs.EOF
			idLocation	= rs("idLocation")
			locName		= rs("locName")
			locState	= rs("locState")
			locTax		= rs("locTax")
			locShipZone	= rs("locShipZone")
			locStatus   = rs("locStatus")
%>
			<tr>
			
				<td valign=top><%=locState%></td>
				<td valign=top><%=locName%></td>
				<td valign=top><%=formatNumber(locTax,2)%></td>
				<td valign=top><%=locShipZone%></td>
				<td valign=top><%=locStatus%></td>
				<td align=right valign=top nowrap>
					[ 
					<a href="SA_loc_edit.asp?action=editstate&recid=<%=idLocation%>">edit</a> | 
					<a href="SA_loc_exec.asp?action=del&idLocation=<%=idLocation%>">delete</a> 
					]
				</td>
				
			</tr>
<%
			rs.MoveNext
		loop
		call closeRS(rs)
%>
	</table>
<%
end if

'Edit State
if action = "editstate" then
%>
	<span class="textBlockHead">Edit State or Province</span>
	<br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td nowrap><b>Code</b></td>
			<td nowrap><b>State/Province Name</b></td>
			<td nowrap><b>Tax %</b></td>
			<td nowrap><b>Zone</b></td>
			<td nowrap><b>Status</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_loc_exec.asp" name="editState">
			<td><input type=text name=locState    id=locState    size=2  maxlength=2   value="<%=locState%>"></td>
			<td><input type=text name=locName     id=locName     size=30 maxlength=100 value="<%=locName%>"></td>
			<td><input type=text name=locTax      id=locTax      size=5  maxlength=10  value="<%=formatNumber(locTax,2)%>"></td>
			<td><input type=text name=locShipZone id=locShipZone size=2  maxlength=2   value="<%=locShipZone%>"></td>
			<td>
				<select name=locStatus id=locStatus size=1>
					<option value="A" <%=checkMatch(locStatus,"A")%>>Active</option>
					<option value="I" <%=checkMatch(locStatus,"I")%>>InActive</option>
				</select>
			</td>
			<td>
				<input type=hidden name=idLocation id=idLocation value="<%=idLocation%>">
				<input type=hidden name=locCountry id=locCountry value="<%=locCountry%>">
				<input type=hidden name=action     id=action     value="editState">
				<input type=submit name=submit1    id=submit1    value="Update">
			</td>
			</form>
		</tr>
	</table>
<%
end if

'Add
if action = "add" then
%>
	<span class="textBlockHead">Add Country</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td nowrap><b>Code</b></td>
			<td nowrap><b>Country Name</b></td>
			<td nowrap><b>Tax %</b></td>
			<td nowrap><b>Zone</b></td>
			<td nowrap><b>Status</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_loc_exec.asp" name="addCountry">
			<td><input type=text name=locCountry  id=locCountry  size=2  maxlength=2></td>
			<td><input type=text name=locName     id=locName     size=30 maxlength=100></td>
			<td><input type=text name=locTax      id=locTax      size=5  maxlength=10 value="0.00"></td>
			<td><input type=text name=locShipZone id=locShipZone size=2  maxlength=2></td>
			<td>
				<select name=locStatus id=locStatus size=1>
					<option value="A">Active</option>
					<option value="I">InActive</option>
				</select>
			</td>
			<td>
				<input type=hidden name=action   id=action   value="add">
				<input type=submit name=submit1  id=submit1  value=" Add ">
			</td>
			</form>
		</tr>
	</table>
<%
end if

'Delete
if action = "del" then
%>
	<span class="textBlockHead">Delete Country</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td nowrap><b>Code</b></td>
			<td nowrap><b>Country Name</b></td>
			<td nowrap><b>Tax %</b></td>
			<td nowrap><b>Zone</b></td>
			<td nowrap><b>Status</b></td>
			<td>&nbsp;</td>
		</tr>
		<form method="post" action="SA_loc_exec.asp" name="deleteCountry">
		<tr>
			<td><%=locCountry%></td>
			<td><%=locName%></td>
			<td><%=formatNumber(locTax,2)%></td>
			<td><%=locShipZone%></td>
			<td><%=locStatus%></td>
			<td>
				<input type=hidden name=idLocation id=idLocation value="<%=idLocation%>">
				<input type=hidden name=action     id=action     value="del">
				<input type=submit name=submit1    id=submit1    value="Delete">
			</td>
		</tr>
		<tr>
			<td valign=middle colspan=6>
				<font color=red>
				Note : All State and/or Province records associated 
				with this Country will also be deleted.
				</font>
			</td>
		</tr>
		</form>
	</table>
<%
end if

if action = "edit" or action = "editstate" or action = "add" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>
	
	<b>Code</b> - Mandatory. This may be any combination of AlphaNumeric 
	characters (maximum 2). Applies to both Country and States/Provinces 
	records. When assigning Codes to a Country, bear in mind that the 
	Country Code must always be unique between Countries. When assigning 
	Codes to a State or Province, the Code must always be unique between 
	States/Provinces within a Country. In other words, it is possible to 
	re-use a State/Province Code as long as it's for another Country (even 
	though the State/Province "Name" will most likely be different). Either 
	way, the application validates the Codes to ensure there are no errors. 
	A good tip is to use the standard Codes for Countries and States where 
	they are available (e.g. "US" for the USA, and "NY" for the state of 
	New York, etc.).<br><br>
	
	<b>Name</b> - Mandatory. In the case of a Country, it will be the name 
	of that particular Country. Similarly, in the case of a State/Province, 
	this will be the name of the State/Province. This is also the name that 
	will be displayed on the checkout forms and orders.<br><br>
	
	<b>Tax %</b> - Mandatory. The value in this field indicates a 
	percentage which must be added to the order total for Sales Tax 
	purposes. If no Tax is to be added, this field must have a value of 
	"0.00". In the event that a Country has State/Province records, the Tax 
	Rate of the State/Province record is always used. If a Country has no 
	State/Province records, then the Country's Tax Rate is applied.<br><br>
	
	<b>Shipping Zone</b> - Mandatory. This must be a numeric value. This 
	field is used to logically group together Countries and/or 
	States/Provinces for the purpose of calculating Shipping. If a Country 
	has States/Provinces, the Shipping Zone of the State/Province will be 
	used, otherwise the Country's Shipping Zone will be used. You may 
	assign a Shipping Zone without creating associated Shipping Rate 
	records for that Shipping Zone. This will however suppress the display 
	of the Country and/or State/Province record(s) when the Customer is 
	checking out. This is to protect you from inadvertantly accepting orders 
	to locations for which you have not defined Shipping Rates.<br><br>
	
	<b>Status</b> - Mandatory. Indicates if a Country or State will 
	be available for selection in your store.<br><br>

	</td></tr>
	</table>
<%
end if

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
	<a href="SA_loc.asp?recallCookie=1">List Countries</a> | 
	<a href="SA_loc_edit.asp?action=edit&recid=<%=idLocation%>">Edit</a> | 
	<a href="SA_loc_edit.asp?action=del&recid=<%=idLocation%>">Delete</a> 
	]
<%
end sub
%>