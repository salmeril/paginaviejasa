<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Shipping Rates Maintenance
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

'ShipRates
dim idShip
dim idShipMethod
dim locShipZone
dim unitType
dim unitsFrom
dim unitsTo
dim addAmt
dim addPerc

'ShipMethod
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

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Shipping - Store Rates Maintenance</font></b>
	<br><br>
</P>

<%
'Page Tabs
call shipTabs("CR")

'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if action <> "edit" and action <> "del" and action <> "add" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idShip
if action = "edit" or action = "del" then
	idShip = trim(Request.QueryString("recId"))
	if len(idShip) = 0 then
		idShip = trim(Request.Form("recId"))
	end if
	if idShip = "" or not isNumeric(idShip) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if
end if

'Get Shipping Rate Record
if action = "edit" or action = "del" then
	mySQL="SELECT * " _
	    & "FROM   shipRates " _
	    & "WHERE  idShip = " & idShip
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	else
		idShipMethod	= rs("idShipMethod")
		locShipZone		= rs("locShipZone")
		unitType		= rs("unitType")
		unitsFrom		= rs("unitsFrom")
		unitsTo			= rs("unitsTo")
		addAmt			= rs("addAmt")
		addPerc			= rs("addPerc")
	end if
	call closeRS(rs)
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
	<span class="textBlockHead">Edit Shipping Rate</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<form method="post" action="SA_shipRate_exec.asp" name="form1">
		<tr>
			<td align=right nowrap><b>Shipping Method</b></td>
			<td align=left>
				<select name=idShipMethod id=idShipMethod size=1>
					<option value="">-- Select --</option>
<%
					mySQL = "SELECT   idShipMethod,shipDesc " _
					      & "FROM     shipMethod " _
					      & "ORDER BY shipDesc"
					set rs = openRSexecute(mySQL)
					do while not rs.EOF
%>
						<option value="<%=rs("idShipMethod")%>" <%=checkMatch(idShipMethod,rs("idShipMethod"))%>><%=rs("shipDesc")%></option>
<%
						rs.MoveNext
					loop
					call closeRS(rs)
%>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Shipping Zone</b></td>
			<td align=left>
				<input type=text name=locShipZone id=locShipZone size=2 maxlength=2 value="<%=locShipZone%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Rate Type</b></td>
			<td align=left>
				<select name=unitType id=unitType size=1>
					<option value="">-- Select --</option>
					<option value="P" <%=checkMatch(unitType,"P")%>>Price</option>
					<option value="W" <%=checkMatch(unitType,"W")%>>Weight</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Price/Weight From</b></td>
			<td align=left>
				<input type=text name=unitsFrom id=unitsFrom size=10 maxlength=10 value="<%=unitsFrom%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Price/Weight To</b></td>
			<td align=left>
				<input type=text name=unitsTo id=unitsTo size=10 maxlength=10 value="<%=unitsTo%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Add Amount</b></td>
			<td align=left>
				<input type=text name=addAmt id=addAmt size=10 maxlength=10 value="<%=addAmt%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Add Percentage</b></td>
			<td align=left>
				<input type=text name=addPerc id=addPerc size=10 maxlength=10 value="<%=addPerc%>">
			</td>
		</tr>
		<tr>
			<td colspan=2 align=center valign=middle>
				<br>
				<input type=hidden name=idShip  id=idShip  value="<%=idShip%>">
				<input type=hidden name=action  id=action  value="edit">
				<input type=submit name=submit1 id=submit1 value="Update Shipping Rate">
			</td>
		</tr>
		</form>
	</table>
<%
end if

'Add
if action = "add" then
%>
	<span class="textBlockHead">Add Shipping Rate</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<form method="post" action="SA_shipRate_exec.asp" name="form1">
		<tr>
			<td align=right nowrap><b>Shipping Method</b></td>
			<td align=left>
				<select name=idShipMethod id=idShipMethod size=1>
					<option value="">-- Select --</option>
<%
					mySQL = "SELECT   idShipMethod,shipDesc " _
					      & "FROM     shipMethod " _
					      & "ORDER BY shipDesc"
					set rs = openRSexecute(mySQL)
					do while not rs.EOF
%>
						<option value="<%=rs("idShipMethod")%>"><%=rs("shipDesc")%></option>
<%
						rs.MoveNext
					loop
					call closeRS(rs)
%>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Shipping Zone</b></td>
			<td align=left>
				<input type=text name=locShipZone id=locShipZone size=2 maxlength=2>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Rate Type</b></td>
			<td align=left>
				<select name=unitType id=unitType size=1>
					<option value="">-- Select --</option>
					<option value="P">Price</option>
					<option value="W">Weight</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Price/Weight From</b></td>
			<td align=left>
				<input type=text name=unitsFrom id=unitsFrom size=10 maxlength=10>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Price/Weight To</b></td>
			<td align=left>
				<input type=text name=unitsTo id=unitsTo size=10 maxlength=10>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Add Amount</b></td>
			<td align=left>
				<input type=text name=addAmt id=addAmt size=10 maxlength=10>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Add Percentage</b></td>
			<td align=left>
				<input type=text name=addPerc id=addPerc size=10 maxlength=10>
			</td>
		</tr>
		<tr>
			<td colspan=2 align=center valign=middle>
				<br>
				<input type=hidden name=action  id=action  value="add">
				<input type=submit name=submit1 id=submit1 value="Add Shipping Rate">
			</td>
		</tr>
		</form>
	</table>
<%
end if

'Delete
if action = "del" then
%>
	<span class="textBlockHead">Delete Shipping Rate</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td align=right nowrap><b>Shipping Method</b></td>
			<td align=left>
<%
				mySQL = "SELECT   shipDesc " _
				      & "FROM     shipMethod " _
				      & "WHERE    idShipMethod = " & idShipMethod
				set rs = openRSexecute(mySQL)
				Response.Write rs("shipDesc")
				call closeRS(rs)
%>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Zone</b></td>
			<td align=left>
				<%=locShipZone%>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Rate Type</b></td>
			<td align=left>
<%
				select case UCase(unitType)
				case "P"
					Response.Write "Price"
				case "W"
					Response.Write "Weight"
				case else
					Response.Write "Unknown"
				end select
%>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Price/Weight From</b></td>
			<td align=left>
				<%=unitsFrom%>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Price/Weight To</b></td>
			<td align=left>
				<%=unitsTo%>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Add Amount</b></td>
			<td align=left>
				<%=addAmt%>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Add Percentage</b></td>
			<td align=left>
				<%=addPerc%>
			</td>
		</tr>
		<tr>
			<form method="post" action="SA_shipRate_exec.asp" name="form4">
			<td colspan=2 align=center valign=middle>
				<br>
				<input type=hidden name=idShip  id=idShip  value="<%=idShip%>">
				<input type=hidden name=action  id=action  value="del">
				<input type=submit name=submit1 id=submit1 value="Delete Shipping Rate">
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
	
		<b>Shipping Method</b> - Mandatory. Select the Shipping Method to 
		which this Shipping Rate record will apply. Shipping Methods are 
		defined and maintained with the <a href="SA_shipMet.asp">Shipping 
		Method Maintenance</a> function.<br><br>
	
		<b>Shipping Zone</b> - Mandatory. Must be a valid numeric value. 
		Enter the Shipping Zone to which this Shipping Rate record will 
		apply. This value is used in conjunction with the Shipping Zone value 
		entered in the <a href="SA_loc.asp?resetCookie=1">Location Maintenance</a> 
		function. When the customer checks out his order, the customer has to 
		enter a shipping address. The system will then compare the Shipping 
		Zone of the cusomer's Country and State (or Province) with the 
		Shipping Zones in the Shipping Rates file to compute Shipping Rates 
		for the customer's shipping address.<br><br>
	
		<b>Rate Type</b> - Mandatory. This value indicates if this Shipping 
		Rate record will be used to calculate shipping based on the Total 
		Order Amount (Total Price exluding Discounts and Taxes) or the 
		Total Weight of the Order (Weight).<br><br>
	
		<b>Price/Weight From and To</b> - Mandatory. These two fields are 
		used to indicate a Price or Weight range (depending on the setting 
		of Rate Type) for this Shipping Rate record. If you selected "Price" 
		as the Rate Type, you will enter a Price range into these fields. 
		If you selected "Weight" as the Rate Type, you will enter a Weight 
		range into these fields.<br>
		<ul>
			<li>You MUST use the same Unit of Weight (Kilogram, Grams, 
			Pounds, etc.) throughout the store. So, if you enter the 
			Weight in "Pounds" for your Products, you must also enter 
			the Weight in "Pounds" here.</li>
			<li>If you specify a price range, note that the shipping rate 
			will be calculated on the gross total order amount, excluding 
			any discounts and taxes.</li>
		</ul>
	
		<b>Add Amount</b> - Optional if a Percentage is entered. This is 
		the amount that will be added to the order for Shipping. You may 
		enter both an Amount as well as a Percentage (the system will 
		calculate both and add the highest value).<br><br>
	
		<b>Add Percentage</b> - Optional if Add Amount is entered. This is 
		the Percentage of the order total which will be added to the order 
		for Shipping. You may enter both an Amount and Percentage (the system 
		will calculate both and add the highest value).<br><br>
	
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
	<a href=SA_shipRate.asp?recallCookie=1>List Rates</a> | 
	<a href=SA_shipRate_edit.asp?action=add>Add</a> | 
	<a href=SA_shipRate_edit.asp?action=edit&recId=<%=idShip%>>Edit</a> | 
	<a href=SA_shipRate_edit.asp?action=del&recId=<%=idShip%>>Delete</a> 
	]
<%
end sub
%>