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
<%

'Database
dim mySQL, cn, rs, rs2

'ShipMethod
dim idShipMethod
dim shipDesc
dim status

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Set Row Colors
dim rowColor, col1, col2
col1 = "#DDDDDD"
col2 = "#EEEEEE"

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Shipping - Store Method Maintenance</font></b>
	<br><br>
</P>

<script language="Javascript">
<!--
	function confirmSubmit()
	{
		var agree=confirm("This action can not be undone. Are you sure you want to continue?");
		if (agree)
			return true ;
		else
			return false ;
	}
	
	function CheckAll(formObject) 
	{
		var chk = formObject.checkAll.checked;
		var len = formObject.elements.length;
		for(var i=0; len >i; i++) 
		{
			var elm = formObject.elements[i];
			if (elm.type == "checkbox")
			{
				elm.checked = chk;
			}
		}
	}
-->
</script>

<%
'Page Tabs
call shipTabs("CM")

if len(trim(Request.QueryString("msg"))) > 0 then
%>
	<font color=red><%=Request.QueryString("msg")%></font>
	<br><br>
<%
end if
%>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="findTable">
	<tr>
		<td align=left valign=middle nowrap>
			<a href="SA_shipMet_edit.asp?action=add">Add Shipping Method</a>
		</td>
	</tr>
</table>

<br>
	
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="listTable">
	<tr>
		<td colspan=5 class="listRowTop">&nbsp;</td>
	</tr>
<%
	rowColor = col1
%>
	<form method="post" action="SA_shipMet_exec.asp" name="form2" id="form2">
	<tr>
		<td class="listRowHead"><b>ID</b></td>
		<td class="listRowHead"><b>Description</b></td>
		<td class="listRowHead"><b>Status</b></td>
		<td class="listRowHead" width="1%"><b>&nbsp;</b></td>
		<td class="listRowHead" align=center width="1%">
			<input type="checkbox" name="checkAll" value="1" onclick="javascript:CheckAll(this.form);">
		</td>
	</tr>
<%
	rowColor = col2

	'Retrieve all Shipping Methods
	mySQL="SELECT * " _
	    & "FROM   ShipMethod " _
	    & "ORDER BY shipDesc"
	set rs = openRSexecute(mySQL)
	do while not rs.eof
%>
		<tr>
			<td bgcolor="<%=rowColor%>" valign=top><%=rs("idShipMethod")%></td>
			<td bgcolor="<%=rowColor%>" valign=top><%=rs("shipDesc")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>" valign=top><%=rs("status")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>" valign=top nowrap>
				[ <a href="SA_shipMet_edit.asp?action=edit&recid=<%=rs("idShipMethod")%>">edit</a> ]
			</td>
			<td align=middle valign=top bgcolor="<%=rowColor%>">
				<input type=checkbox name="idShipMethod" id="idShipMethod" value="<%=rs("idShipMethod")%>">
			</td>
		</tr>
<%
		rs.movenext
		
		'Switch Row Color
		if rowColor = col2 then
			rowColor = col1
		else
			rowColor = col2
		end if

	loop
%>
	<tr>
		<td colspan=5 align=right nowrap class="listRowBot">
			<input type=hidden name="action" id="action" value="bulkDel">
			Delete Selected Shipping Methods? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
		</td>
	</tr>
	</form>
<%
	call closeRS(rs)
%>
</table>

<br>

<span class="textBlockHead">Help and Instructions :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr><td>

	<b>Overview</b> - Store Shipping Methods are used to group 
	together a common set of Store Shipping Rates. Store Methods 
	and Rates are defined by you, and will therefore be stored and 
	retrieved directly from your database. For example, you can 
	create a Store Shipping Method called 'FedEx Ground', and then 
	create your own unique set of Store Shipping Rates for the 'FedEx 
	Ground' Shipping Method.<br><br>
	
	A Shipping Method will also be applied to one or more Shipping 
	Zones (Zones are set up using the 'Locations Maintenance' 
	function). Each Shipping Method/Zone combination will have it's 
	own unique set of Shipping Rates.<br><br>
	
	<b>Example :</b>
	<br>
	<ul>
		<li><font color=blue>FedEx Ground</font>
			<ul>
				<li>USA & Canada (Zone 1)
					<ul>
						<li>Shipping Rates based on total Order Amount
							<ul>
								<li>$000.00 to $100.00 -> add $10.00 to order</li>
							</ul>
							<ul>
								<li>$100.01 to $200.00 -> add $15.00 to order</li>
							</ul>
							<ul>
								<li>$200.01 to $500.00 -> add $20.00 to order</li>
							</ul>
							<ul>
								<li>etc.</li>
							</ul>
						</li>
						<li>Shipping Rates based on Order Weight (Pounds, Kilograms, etc.)
							<ul>
								<li>00 to 10 -> add $12.50 to order</li>
							</ul>
							<ul>
								<li>11 to 20 -> add $14.50 to order</li>
							</ul>
							<ul>
								<li>21 to 30 -> add $17.00 to order</li>
							</ul>
							<ul>
								<li>etc.</li>
							</ul>
						</li>
					</ul>
				</li>
				<li>Europe (Zone 2)
					<ul>
						<li>Shipping Rates based on Order Weight (Pounds, Kilograms, etc.)
							<ul>
								<li>00 to 10 -> add 10.00% to order</li>
							</ul>
							<ul>
								<li>11 to 20 -> add 12.50% to order</li>
							</ul>
							<ul>
								<li>21 to 30 -> add 15.00% to order</li>
							</ul>
							<ul>
								<li>etc.</li>
							</ul>
						</li>
					</ul>
				</li>
			</ul>
		</li>
	</ul>
		
</td></tr>
</table>
<%
call closedb()
%>
<!--#include file="_INCfooter_.asp"-->