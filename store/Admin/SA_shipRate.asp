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
dim mySQL, cn, rs, rs2

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
dim I
dim item
dim count
dim pageSize
dim totalPages
dim showArr

dim curPage
dim showShipMet
dim showShipZone
dim showType

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
pageSize = 50

'Get Page to show
curPage = Request.Form("curPage")						'Form
if len(curPage) = 0 then
	curPage = Request.QueryString("curPage")			'QueryString
end if

'Get showShipMet
showShipMet = Request.Form("showShipMet")				'Form
if len(showShipMet) = 0 then
	showShipMet = Request.QueryString("showShipMet")	'QueryString
end if

'Get showShipZone
showShipZone = Request.Form("showShipZone")				'Form
if len(showShipZone) = 0 then
	showShipZone = Request.QueryString("showShipZone")	'QueryString
end if

'Get showType
showType = Request.Form("showType")						'Form
if len(showType) = 0 then
	showType = Request.QueryString("showType")			'QueryString
end if

'Check if a Cookie Reset was requested
if Request.QueryString("resetCookie") = "1" then
	Response.Cookies("ShipSearch").expires = Date() - 30
else
	'Check if a Cookie Recall was requested
	if Request.QueryString("recallCookie") = "1" then
		for each item in Request.Cookies
			if item = "ShipSearch" then
				showArr		= Split(Request.Cookies(item),"*|*")
				curPage		= showArr(0)
				showShipMet	= showArr(1)
				showShipZone= showArr(2)
				showType	= showArr(3)
			end if
		next
	else
		'Save Search Criteria in a Cookie
		Response.Cookies("ShipSearch") = navCookie(curPage)
		Response.Cookies("ShipSearch").expires = Date() + 30
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
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Shipping - Store Rates Maintenance</font></b>
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
call shipTabs("CR")

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
			<a href="SA_shipRate_edit.asp?action=add">Add Shipping Rate</a>
		</td>
		
		<form method="post" action="SA_shipRate.asp" name="form4">
		<td align=right valign=top nowrap>
			<select name=showShipMet id=showShipMet size=1>
				<option value="" <%=checkMatch(showShipMet,"")%>>Shipping Method</option>
<%
				mySQL = "SELECT   idShipMethod,shipDesc " _
				      & "FROM     shipMethod " _
				      & "ORDER BY shipDesc"
				set rs = openRSexecute(mySQL)
				do while not rs.EOF
					Response.Write "<option value=""" _
								 & rs("idShipMethod") _
								 & """ " _
								 & checkMatch(showShipMet,rs("idShipMethod")) _
								 & ">" _
								 & rs("shipDesc") _
								 & "</option>"
					rs.MoveNext
				loop
				call closeRS(rs)
%>
			</select>&nbsp;
			<select name=showShipZone id=showShipZone size=1>
				<option value="" <%=checkMatch(showShipZone,"")%>>Zone</option>
<%
				mySQL = "SELECT   locShipZone " _
				      & "FROM     locations " _
				      & "GROUP BY locShipZone " _
				      & "ORDER BY locShipZone "
				set rs = openRSexecute(mySQL)
				do while not rs.EOF
					Response.Write "<option value=""" _
								 & rs("locShipZone") _
								 & """ " _
								 & checkMatch(showShipZone,rs("locShipZone")) _
								 & ">" _
								 & rs("locShipZone") _
								 & "</option>"
					rs.MoveNext
				loop
				call closeRS(rs)
%>
			</select>&nbsp;
			<select name=showType id=showType size=1>
				<option value=""  <%=checkMatch(showType,"") %>>Rate Type</option>
				<option value="P" <%=checkMatch(showType,"P")%>>Price</option>
				<option value="W" <%=checkMatch(showType,"W")%>>Weight</option>
			</select>&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
		
	</tr>
	
</table>

<br>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="listTable">
<%
	'Specify fields and table
	mySQL="SELECT idShip,shipRates.idShipMethod,locShipZone," _
		& "       unitType,unitsFrom,unitsTo,addAmt,addPerc, " _
	    & "       shipMethod.shipDesc " _
	    & "FROM   shipRates, shipMethod " _
	    & "WHERE  shipRates.idShipMethod = shipMethod.idShipMethod " 
	    
	'Shipping Method
	if len(showShipMet) > 0 then
		mySQL = mySQL & "AND shipRates.idShipMethod = " & showShipMet & " "
	end if
	    
	'Shipping Zone
	if len(showShipZone) > 0 then
		mySQL = mySQL & "AND locShipZone = " & showShipZone & " "
	end if
	    
	'Shipping Rate Type
	if len(showType) > 0 then
		mySQL = mySQL & "AND unitType = '" & showType & "' "
	end if
	
	'Sort Order
	mySQL = mySQL & "ORDER BY shipDesc,locShipZone,unitType," _
				  & "         unitsFrom,unitsTo"
	    
	set rs = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pageSize)
	if rs.eof then
%>
		<tr>
			<td align=center valign=middle>
				<br>
				<b>No Shipping Rates matched search criteria.</b>
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
			<td colspan=9 class="listRowTop">
<%
				call pageNavigation("selectPageTop")
%>
			</td>
		</tr>

		<form method="post" action="SA_shipRate_exec.asp" name="form2" id="form2">
		<tr>
			<td class="listRowHead"><b>Shipping Method</b></td>
			<td class="listRowHead"><b>Zone</b></td>
			<td class="listRowHead"><b>Type</b></td>
			<td class="listRowHead" align=right><b>From</b></td>
			<td class="listRowHead" align=right><b>To</b></td>
			<td class="listRowHead" align=right><b>Amount</b></td>
			<td class="listRowHead" align=right><b>Perc.</b></td>
			<td class="listRowHead"><b>&nbsp;</b></td>
			<td class="listRowHead" nowrap align=center>
				<input type="checkbox" name="checkAll" value="1" onclick="javascript:CheckAll(this.form);">
			</td>
		</tr>
<%
		rowColor = col2
		do while not rs.eof and count < rs.pageSize
%>
			<tr>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("shipDesc")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("locShipZone")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top>
<%
					select case UCase(rs("unitType"))
					case "P"
						Response.Write "Price"
					case "W"
						Response.Write "Weight"
					case else
						Response.Write "Unknown"
					end select
%>
				</td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap><%=rs("unitsFrom")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap><%=rs("unitsTo")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap>
<%
					if isNull(rs("addAmt")) then
						Response.Write "-"
					else
						Response.Write moneyD(rs("addAmt"))
					end if
%>
				</td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap>
<%
					if isNull(rs("addPerc")) then
						Response.Write "-"
					else
						Response.Write rs("addPerc") & "%"
					end if
%>
				</td>
				<td bgcolor="<%=rowColor%>" align=right valign=top nowrap>
					[ <a href="SA_shipRate_edit.asp?action=edit&recid=<%=rs("idShip")%>">edit</a> ]
				</td>
				<td align=middle valign=top bgcolor="<%=rowColor%>">
					<input type=checkbox name="idShip" id="idShip" value="<%=rs("idShip")%>">
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
			<td nowrap colspan=5 align=right class="listRowBot">
				<input type=hidden name="action" id="action" value="bulkDel">
				Delete Selected Shipping Rates? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
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
	
	<b>Overview</b> - Store Shipping Rates are specified for each 
	unique Store Shipping Method and Zone combination. Shipping Rates 
	are applied to an order based on the order's total Weight and/or 
	Price range. This allows for a tremendous amount of flexibility. 
	You can also specify if the shipping rate for the order must be 
	calculated as a fixed amount, or a percentage of the order total.<br><br>
	
	It is also possible to specify both a fixed amount AND a percentage 
	for a specific price/weight range. In these cases, the system will 
	display the higher value of the two calculations.<br><br>
	
	<b>Example</b>
	<br>
	<ul>
		<li>FedEx Ground (Shipping Method)
			<ul>
				<li>USA & Canada (Zone 1)
					<ul>
						<li><font color=blue>Shipping Rates based on total Order Amount</font>
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
						<li><font color=blue>Shipping Rates based on Order Weight (Pounds, Kilograms, etc.)</font>
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
						<li><font color=blue>Shipping Rates based on Order Weight (Pounds, Kilograms, etc.)</font>
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
<%
'*********************************************************************
'Make QueryString for Paging
'*********************************************************************
function navQueryStr(pageNum)

	navQueryStr = "?curPage="		& server.URLEncode(pageNum) _
	            & "&showShipMet="	& server.URLEncode(showShipMet) _
	            & "&showShipZone="	& server.URLEncode(showShipZone) _
	            & "&showType="		& server.URLEncode(showType)
end function
'*********************************************************************
'Make Cookie Value for Paging
'*********************************************************************
function navCookie(pageNum)

	navCookie = pageNum		 & "*|*" _
	          & showShipMet	 & "*|*" _
	          & showShipZone & "*|*" _
	          & showType
end function
'*********************************************************************
'Display page navigation
'*********************************************************************
sub pageNavigation(formFieldName)
	Response.Write "Page "
	Response.Write "<select onChange=""location.href=this.options[selectedIndex].value"" name=" & trim(formFieldName) & ">"
	for I = 1 to TotalPages
		Response.Write "<option value=""SA_shipRate.asp" & navQueryStr(I) & """ " & checkMatch(curPage,I) & ">" & I & "</option>" & vbCrlf
	next
	Response.Write "</select>&nbsp;of&nbsp;" & TotalPages & "&nbsp;&nbsp;"
	Response.Write "[&nbsp;"
	if curPage > 1 then
		Response.Write "<a href=""SA_shipRate.asp" & navQueryStr(curPage-1) & """>Back</a>"
	else
		Response.Write "Back"
	end if
	Response.Write "&nbsp;|&nbsp;"
	if curPage < TotalPages then
		Response.Write "<a href=""SA_shipRate.asp" & navQueryStr(curPage+1) & """>Next</a>"
	else
		Response.Write "Next"
	end if
	Response.Write "&nbsp;]"
end sub
%>