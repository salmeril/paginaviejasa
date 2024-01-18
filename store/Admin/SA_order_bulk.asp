<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Bulk Order Maintenance
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
<%

'Database
dim mySQL, cn, rs

'Products
dim idProduct
dim description
dim price
dim sku
dim weight

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

%>
<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Bulk Order Maintenance</font></b>
	<br><br>
</P>

<%
if len(trim(Request.QueryString("msg"))) > 0 then
%>
	<font color=red><%=Request.QueryString("msg")%></font><br><br>
<%
end if
%>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr><td>

	<font color=red size=1>WARNING :</font> These functions should 
	be used with caution. Once an update has been applied, it can 
	not be undone easily. These functions will not be appropriate 
	for everyone, but we decided to include them because we use 
	them, and thought others may want to use them as well.
	
</td></tr>
</table>

<br>

<span class="textBlockHead">ADD Items to Orders</span><br>
<%
'Get Product Records
mySQL="SELECT idProduct,description,price,sku,weight " _
    & "FROM   products " _
    & "ORDER BY description"
set rs = openRSexecute(mySQL)
if rs.EOF then
	Response.Clear
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("At least one Product is required in the Database to use this function.")
end if
%>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr>
		<form method="post" action="SA_order_bulk_exec.asp" name="form1">
		<td>
			1. Search for all Orders with this Item :<br>
			<select name=idProduct1 id=idProduct1 size=1>
				<option value=""></option>
<%
				rs.moveFirst
				do while not rs.eof
					Response.Write "<option value=""" & rs("idProduct") & """>" & rs("description") & " (" & rs("sku") &  ")</option>"
					rs.movenext
				loop
%>
			</select><br><br>
			2. Add this Item to the Order :<br>
			<select name=idProduct2 id=idProduct2 size=1>
				<option value=""></option>
<%
				rs.moveFirst
				do while not rs.eof
					Response.Write "<option value=""" & rs("idProduct") & """>" & rs("description") & " (" & rs("sku") &  ")</option>"
					rs.movenext
				loop
%>
			</select><br>
			<input type=checkbox name=invertPrice value="Y"> Invert the Price (ie. Price x -1).<br><br>
			
			<input type=hidden name=action   id=action    value="AddItem">
			<input type=submit name=submit1  id=submit1   value="Update"><br><br>
			
			<b>1.</b> SKU, Description, Price and Weight for the 
			added item are obtained from the Product file.<br>
			<b>2.</b> Quantity will always be 1.<br>
			<b>3.</b> Duplicates are NOT checked. If the added Item 
			already exists on the order, it will be re-added.<br>
			<b>4.</b> EXISTING order information WILL NOT BE CHANGED, 
			including the Order Status and Order Totals, which may 
			result in the individual items not adding up to the Order 
			Total. This is to prevent possible inconsistencies between 
			the Order Total and any actual amounts already paid. If 
			you want the individual Order Items to add up to the Order 
			Total, add the new item twice, once with a positive price, 
			then with a negative price. Or you could simply make the 
			price of the product you are adding "0.00".<br>
			
			<br>

		</td>
		</form>
	</tr>
</table>
<%
'Close Recordset
call closeRS(rs)

'Close Database Connection
call closedb()
%>
<!--#include file="_INCfooter_.asp"-->