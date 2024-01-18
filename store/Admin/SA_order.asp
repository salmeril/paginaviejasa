<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Order Maintenance
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

'cartHead
dim idOrder
dim idCust
dim orderDate
dim orderDateInt
dim randomKey
dim subTotal
dim taxTotal
dim shipmentTotal
dim Total
dim shipmentMethod
dim name
dim lastName
dim customerCompany
dim phone
dim email
dim address
dim city
dim locState
dim locCountry
dim zip
dim shippingName
dim shippingLastName
dim shippingAddress
dim shippingCity
dim shippingLocState
dim shippingLocCountry
dim shippingZip
dim paymentType
dim cardType
dim cardNumber
dim cardExpMonth
dim cardExpYear
dim cardVerify
dim cardName
dim generalComments
dim orderStatus
dim auditInfo
dim storeComments
dim storeCommentsPriv

'cartRows
dim idCartRow
dim idProduct
dim sku
dim quantity
dim unitPrice
dim unitWeight
dim description

'CartRowsOptions
dim idCartRowOption
dim idOption
dim optionPrice
dim optionDescrip

'Work Fields
dim I
dim item
dim count
dim pageSize
dim totalPages
dim showArr
dim sortField
dim delUordHours

dim curPage
dim showStatus
dim showField
dim showPhrase

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

'Set Number of Items per Page
pageSize = 20

'Get Page to show
curPage = Request.Form("curPage")						'Form
if len(curPage) = 0 then
	curPage = Request.QueryString("curPage")			'QueryString
end if

'Get showStatus
showStatus = Request.Form("showStatus")					'Form
if len(showStatus) = 0 then
	showStatus = Request.QueryString("showStatus")		'QueryString
end if

'Get showField
showField = Request.Form("showField")					'Form
if len(showField) = 0 then
	showField = Request.QueryString("showField")		'QueryString
end if

'Get showPhrase
showPhrase = Request.Form("showPhrase")					'Form
if len(showPhrase) = 0 then
	showPhrase = Request.QueryString("showPhrase")		'QueryString
end if

'Check if a Cookie Reset was requested
if Request.QueryString("resetCookie") = "1" then
	Response.Cookies("OrderSearch").expires = Date() - 30
else
	'Check if a Cookie Recall was requested
	if Request.QueryString("recallCookie") = "1" then
		for each item in Request.Cookies
			if item = "OrderSearch" then
				showArr		= Split(Request.Cookies(item),"*|*")
				curPage		= showArr(0)
				showStatus	= showArr(1)
				showField	= showArr(2)
				showPhrase	= showArr(3)
			end if
		next
	else
		'Save Search Criteria in a Cookie
		Response.Cookies("OrderSearch") = navCookie(curPage)
		Response.Cookies("OrderSearch").expires = Date() + 30
	end if
end if

'After attempting to retrieve the search criteria through the various 
'mechanisms above (Form/QueryString/Cookie), check that some of the 
'critical values are valid. If not, set to default values.
if len(curPage) = 0 or not isNumeric(curPage) then
	curPage = 1
else
	curPage = CLng(curPage)
end if

'Check what we will be sorting the results on
sortField = Request.Form("sortField")					'Form
if len(sortField) = 0 then
	sortField = Request.QueryString("sortField")		'QueryString
end if
if len(sortField) = 0 then
	sortField = "idOrder DESC"
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Order Maintenance</font></b>
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
			<a href="SA_order_bulk.asp">Bulk Modifications</a>
		</td>

		<form method="post" action="SA_order.asp" name="form4">
		<td align=right valign=top nowrap>
			Show Orders where Status is&nbsp;
			<select name=showStatus id=showStatus size=1>
				<option value="">Show all Orders</option>
				<option value="">-----------------------</option>
				<option value="U" <%=checkMatch(showStatus,"U")%>>Unfinalized</option>
				<option value="S" <%=checkMatch(showStatus,"S")%>>Saved</option>
				<option value="0" <%=checkMatch(showStatus,"0")%>>Pending</option>
				<option value="1" <%=checkMatch(showStatus,"1")%>>Paid</option>
				<option value="2" <%=checkMatch(showStatus,"2")%>>Shipped</option>
				<option value="7" <%=checkMatch(showStatus,"7")%>>Complete</option>
				<option value="9" <%=checkMatch(showStatus,"9")%>>Cancelled</option>
			</select>&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
	</tr>
	<tr>
		<form method="post" action="SA_order.asp" name="form2">
		<td align=right valign=top colspan=2 nowrap>
			Show Orders where&nbsp;
			<select name=showField id=showField size=1>
				<option value="">-- Select --</option>
				<option value="name"        <%=checkMatch(showField,"name")       %>>First Name</option>
				<option value="lastName"    <%=checkMatch(showField,"lastName")   %>>Last Name</option>
				<option value="address"     <%=checkMatch(showField,"address")    %>>Address</option>
				<option value="email"       <%=checkMatch(showField,"email")      %>>Email</option>
				<option value="idCust"      <%=checkMatch(showField,"idCust")     %>>Customer ID</option>
				<option value="sku"         <%=checkMatch(showField,"sku")        %>>Product SKU</option>
				<option value="description" <%=checkMatch(showField,"description")%>>Product Desc.</option>
			</select>&nbsp;
			contains the phrase&nbsp;
			<input type=text name=showPhrase id=showPhrase size=20 maxlength=50 value="<%=showPhrase%>">&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
	</tr>
	<tr>
		<form method="post" action="SA_order_exec.asp" name="delUnfinal">
		<td align=right valign=top colspan=2 nowrap>
			Find and <span style="COLOR: red">DELETE</span> Unfinalized 
			orders older than 
			<select name=delUordHours id=delUordHours size=1>
				<option value="">-- Select --</option>
				<option value="1"  <%=checkMatch(delUordHours,"1") %>>1 Hour</option>
				<option value="2"  <%=checkMatch(delUordHours,"2") %>>2 Hours</option>
				<option value="6"  <%=checkMatch(delUordHours,"6") %>>6 Hours</option>
				<option value="12" <%=checkMatch(delUordHours,"12")%>>12 Hours</option>
				<option value="24" <%=checkMatch(delUordHours,"24")%>>24 Hours</option>
			</select> 
			<input type=hidden name="action" id="action" value="delUord">
			<input type=submit name=submit1  id=submit1  value="Find">
		</td>
		</form>
	</tr>
</table>

<br>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="listTable">
<%
	'Specify fields and table
	mySQL="SELECT idOrder,idCust,orderDate,total,name,lastName," _
		& "       address,paymentType,orderStatus " _
	    & "FROM   cartHead " _
	    & "WHERE  1=1 " 'Dummy check to set up conditional checks below
	    
	'Status
	if len(showStatus) > 0 then
		mySQL = mySQL & "AND orderStatus = '" & showStatus & "' "
	end if

	'Field and Search Phrase
	if len(showField) > 0 and len(showPhrase) > 0 then	
		'If we are searching on idCust, look for exact match
		if lCase(showfield) = "idcust" then
			if IsNumeric(showPhrase) then
				mySQL = mySQL & "AND idCust = " & showPhrase & " "
			else 'Force Database to return no entries
				mySQL = mySQL & "AND 1=2 "
			end if
		else
			'If we are searching on SKU or Product Description, we need 
			'to check against the cartRows file, else we check against 
			'the cartHead file.
			if lCase(showfield) = "sku" or lCase(showField) = "description" then
				mySQL = mySQL & "AND EXISTS (SELECT idCartRow FROM cartRows WHERE cartRows.idOrder = cartHead.idOrder AND " & showField & " LIKE '%" & replace(showPhrase,"'","''") & "%') "
			else
				mySQL = mySQL & "AND " & showField & " LIKE '%" & replace(showPhrase,"'","''") & "%' "
			end if
		end if
	end if

	'Sort Order
	mySQL = mySQL & "ORDER BY " & sortField
	
	set rs = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pageSize)
	if rs.eof then
%>
		<tr>
			<td align=center valign=middle>
				<br>
				<b>No Orders matched search criteria.</b>
				<br><br>
			</td>
		</tr>
<%
	else
		rs.MoveFirst
		rs.PageSize		= pageSize
		totalPages 		= rs.PageCount
		rs.AbsolutePage	= curPage
%>
		<tr>
			<td colspan=4 class="listRowTop">
<%
				call pageNavigation("selectPageTop")
%>
			</td>
			<td colspan=4 align=right class="listRowTop">
				Sort : 
				<select name=sortField id=sortField size=1 onChange="location.href='SA_order.asp?recallCookie=1&sortField='+this.options[selectedIndex].value">
					<option value="idOrder DESC" <%=checkMatch(sortField,"idOrder DESC")%>>Order (Descending)</option>
					<option value="idOrder"      <%=checkMatch(sortField,"idOrder")     %>>Order (Ascending)</option>
					<option value="lastName"     <%=checkMatch(sortField,"lastName")    %>>Last Name</option>
					<option value="paymentType"  <%=checkMatch(sortField,"paymentType") %>>Payment Type</option>
				</select>
			</td>
		</tr>
<%
		rowColor = col1
%>
		<form method="post" action="SA_order_exec.asp" name="form3" id="form3">
		<tr>
			<td class="listRowHead" nowrap><b>Order</b></td>
			<td class="listRowHead" nowrap><b>Date/Time</b></td>
			<td class="listRowHead" nowrap><b>Total</b></td>
			<td class="listRowHead" nowrap><b>Name/Address</b></td>
			<td class="listRowHead" nowrap><b>PayType</b></td>
			<td class="listRowHead" nowrap><b>Status</b></td>
			<td class="listRowHead" nowrap>&nbsp;</td>
			<td class="listRowHead" nowrap align=center>
				<input type="checkbox" name="checkAll" value="1" onclick="javascript:CheckAll(this.form);">
			</td>
		</tr>
<%
		rowColor = col2
		do while not rs.eof and count < rs.pageSize
%>
			<tr>
				<td valign=top nowrap bgcolor="<%=rowColor%>">
					<a href="SA_order_edit.asp?action=view&recid=<%=rs("idOrder")%>"><%=pOrderPrefix & "-" & rs("idOrder")%></a>
				</td>
				<td valign=top nowrap bgcolor="<%=rowColor%>">
					<%=formatTheDate(rs("orderDate"))%><br>
					<%=formatDateTime(rs("orderDate"),3)%>
				</td>
				<td valign=top nowrap bgcolor="<%=rowColor%>">
<%
					if isNull(rs("total")) then
						Response.Write "-"
					else
						Response.Write moneyD(rs("total"))
					end if
%>
				</td>
				<td valign=top bgcolor="<%=rowColor%>">
<%
					if isnull(rs("lastName")) or len(rs("lastName")) = 0 then
						Response.Write "-"
					else
						Response.Write "<a href=""SA_cust_edit.asp?action=edit&recid=" & rs("idCust") & """>" & rs("lastName") & ", " & rs("name") & "</a>"
					end if
					Response.Write "<br>" & rs("address")
%>
				</td>
				<td valign=top nowrap bgcolor="<%=rowColor%>">
<%
					if isnull(rs("paymentType")) or len(rs("paymentType")) = 0 then
						Response.Write "-"
					else
						Response.Write rs("paymentType")
					end if
%>
				</td>
				<td valign=top nowrap bgcolor="<%=rowColor%>">
<%
					Response.Write orderStatusDesc(rs("orderStatus"))
					if pAuthNet = -1 then
						select case rs("orderStatus")
						case "0", "9"
							Response.Write "<br><a href=""SA_authnet.asp?recid=" & rs("idOrder") & """>authorize</a>"
						end select
					end if
%>
				</td>
				<td align=right valign=top nowrap bgcolor="<%=rowColor%>">
					[ 
					<a href="SA_order_edit.asp?action=inv&recid=<%=rs("idOrder")%>" target="_blank">inv</a> | 
					<a href="SA_order_edit.asp?action=edit&recid=<%=rs("idOrder")%>">edit</a> 
					]
				</td>
				<td align=middle valign=top bgcolor="<%=rowColor%>">
					<input type=checkbox name="idOrder" id="idOrder" value="<%=rs("idOrder")%>">
				</td>
			</tr>
<%
			count = count + 1  
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
			<td nowrap colspan=4 class="listRowBot">
<%
				call pageNavigation("selectPageBot")
%>
			</td>
			<td nowrap colspan=4 align=right class="listRowBot">
				<input type=hidden name="action" id="action" value="bulkDel">
				Delete Selected Orders? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
			</td>
		</tr>
		</form>
<%
	end if
	call closeRS(rs)
%>
</table>

<br>

<span class="textBlockHead">Help and Instructions :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr><td>

	<b>Find Orders(s)</b> - You have several options for finding a 
	specific Order. Orders are displayed in Order Number sequence, with 
	the newest orders listed first.<br><br>
	
	1. You can list all Orders which contain a specific phrase in one 
	of several Order fields.<br><br>
	
	2. You can list all Orders with a specified Order Status.<br><br>
	
	<b>Authorize Payment</b> - Click to request payment authorization from 
	Authorize.Net. This option will only be available if you "activate" 
	Authorize.Net in your store's configurations. You will then be able to 
	authorize a Credit Card payment for any order, even if this was not 
	the customer's original prefered method of payment. This way, should 
	the customer change their mind about how they wish to pay for the 
	order, you can collect their Credit Card information over the phone 
	(or fax, email, etc.), authorize a Credit Card payment for them and 
	update the order status to "Paid".<br><br>
	
	<b>Invoice</b> - Click to view a printable Invoice for the Order.<br><br>
	
	<b>Edit Order</b> - Click to change Order Status and/or Information.<br><br>
	
	<b>Delete Order</b> - Check the box next to order(s) you want to 
	delete, and click "Yes".<br><br>
	
	<b>View Order</b> - Click the Order Number to view an Order.
	
</td></tr>
</table>
<%
call closedb()
%>

<!--#include file="_INCfooter_.asp"-->

<%
'*********************************************************************
'Make QueryString for Paging
'*********************************************************************
function navQueryStr(pageNum)

	navQueryStr = "?curPage="		& server.URLEncode(pageNum) _
	            & "&showStatus="	& server.URLEncode(showStatus) _
	            & "&showField="		& server.URLEncode(showField) _
	            & "&showPhrase="	& server.URLEncode(showPhrase)
end function
'*********************************************************************
'Make Cookie Value for Paging
'*********************************************************************
function navCookie(pageNum)

	navCookie = pageNum		& "*|*" _
	          & showStatus	& "*|*" _
	          & showField	& "*|*" _
	          & showPhrase
end function
'*********************************************************************
'Display page navigation
'*********************************************************************
sub pageNavigation(formFieldName)
	Response.Write "Page "
	Response.Write "<select onChange=""location.href=this.options[selectedIndex].value"" name=" & trim(formFieldName) & ">"
	for I = 1 to TotalPages
		Response.Write "<option value=""SA_order.asp" & navQueryStr(I) & "&sortField=" & server.URLEncode(sortField) & """ " & checkMatch(curPage,I) & ">" & I & "</option>" & vbCrlf
	next
	Response.Write "</select>&nbsp;of&nbsp;" & TotalPages & "&nbsp;&nbsp;"
	Response.Write "[&nbsp;"
	if curPage > 1 then
		Response.Write "<a href=""SA_order.asp" & navQueryStr(curPage-1) & "&sortField=" & server.URLEncode(sortField) & """>Back</a>"
	else
		Response.Write "Back"
	end if
	Response.Write "&nbsp;|&nbsp;"
	if curPage < TotalPages then
		Response.Write "<a href=""SA_order.asp" & navQueryStr(curPage+1) & "&sortField=" & server.URLEncode(sortField) & """>Next</a>"
	else
		Response.Write "Next"
	end if
	Response.Write "&nbsp;]"
end sub
%>