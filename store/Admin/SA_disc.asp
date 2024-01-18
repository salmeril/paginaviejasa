<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Order Discount Maintenance
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
dim mySQL, cn, rs, rs2

'DiscOrder
dim idDiscOrder
dim discCode
dim discPerc
dim discAmt
dim discFromAmt
dim discToAmt
dim discStatus
dim discOnceOnly
dim discValidFrom
dim discValidTo

'Work Fields
dim I
dim item
dim count
dim pageSize
dim totalPages
dim showArr
dim sortField

dim curPage
dim showStatus
dim showOnceOnly
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
pageSize = 25

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

'Get showOnceOnly
showOnceOnly = Request.Form("showOnceOnly")				'Form
if len(showOnceOnly) = 0 then
	showOnceOnly = Request.QueryString("showOnceOnly")	'QueryString
end if

'Get showPhrase
showPhrase = Request.Form("showPhrase")					'Form
if len(showPhrase) = 0 then
	showPhrase = Request.QueryString("showPhrase")		'QueryString
end if

'Check if a Cookie Reset was requested
if Request.QueryString("resetCookie") = "1" then
	Response.Cookies("DiscSearch").expires = Date() - 30
else
	'Check if a Cookie Recall was requested
	if Request.QueryString("recallCookie") = "1" then
		for each item in Request.Cookies
			if item = "DiscSearch" then
				showArr			= Split(Request.Cookies(item),"*|*")
				curPage			= showArr(0)
				showStatus		= showArr(1)
				showOnceOnly	= showArr(2)
				showPhrase      = showArr(3)
			end if
		next
	else
		'Save Search Criteria in a Cookie
		Response.Cookies("DiscSearch") = navCookie(curPage)
		Response.Cookies("DiscSearch").expires = Date() + 30
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
	sortField = "discCode"
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Order Discount Maintenance</font></b>
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
			<a href="SA_disc_edit.asp?action=add">Add New Discount</a>
		</td>

		<form method="post" action="SA_disc.asp" name="form4">
		<td align=right valign=top nowrap>
			Status = 
			<select name=showStatus id=showStatus size=1>
				<option value=""  <%=checkMatch(showStatus,"") %>>All</option>
				<option value="A" <%=checkMatch(showStatus,"A")%>>Active</option>
				<option value="I" <%=checkMatch(showStatus,"I")%>>InActive</option>
				<option value="U" <%=checkMatch(showStatus,"U")%>>Used</option>
			</select>&nbsp;
			Once Only = 
			<select name=showOnceOnly id=showOnceOnly size=1>
				<option value=""  <%=checkMatch(showOnceOnly,"") %>>All</option>
				<option value="Y" <%=checkMatch(showOnceOnly,"Y")%>>Yes</option>
				<option value="N" <%=checkMatch(showOnceOnly,"N")%>>No</option>
			</select>&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
	</tr>
	<tr>
		<form method="post" action="SA_disc.asp" name="form2">
		<td colspan=2 align=right valign=top nowrap>
			Show Discounts where Code contains the phrase 
			<input type=text name=showPhrase id=showPhrase size=20 maxlength=50 value="<%=showPhrase%>">&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
	</tr>
</table>

<br>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="listTable">
<%
	'Specify fields and table
	mySQL="SELECT idDiscOrder,discCode,discPerc,discAmt," _
		& "       discFromAmt,discToAmt,discStatus," _
	    & "       discOnceOnly,discValidFrom,discValidTo " _
	    & "FROM   DiscOrder " _
	    & "WHERE  1=1 " 'Dummy check to set up conditional checks below
	    
	'discStatus
	if len(showStatus) > 0 then
		mySQL = mySQL & "AND discStatus = '" & showStatus & "' "
	end if
	    
	'discOnceOnly
	if len(showOnceOnly) > 0 then
		mySQL = mySQL & "AND discOnceOnly = '" & showOnceOnly & "' "
	end if
	
	'Search Phrase
	if len(showPhrase) > 0 then	
		mySQL = mySQL & "AND discCode LIKE '%" & replace(showPhrase,"'","''") & "%' "
	end if
	
	'Sort Order
	mySQL = mySQL & "ORDER BY " & sortField
	
	set rs = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pageSize)
	if rs.eof then
%>
		<tr>
			<td align=center valign=middle>
				<br>
				<b>No Discounts matched search criteria.</b>
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
			<td colspan=5 class="listRowTop">
<%
				call pageNavigation("selectPageTop")
%>
			</td>
			<td colspan=5 align=right class="listRowTop">
				Sort : 
				<select name=sortField id=sortField size=1 onChange="location.href='SA_disc.asp?recallCookie=1&sortField='+this.options[selectedIndex].value">
					<option value="discCode"      <%=checkMatch(sortField,"discCode")     %>>Discount Code</option>
					<option value="discValidFrom" <%=checkMatch(sortField,"discValidFrom")%>>Valid From Date</option>
				</select>
			</td>
		</tr>
<%
		rowColor = col1
%>
		<form method="post" action="SA_disc_exec.asp" name="form3" id="form3">
		<tr>
			<td class="listRowHead"><b>Code</b></td>
			<td class="listRowHead" align=right><b>From</b></td>
			<td class="listRowHead" align=right><b>To</b></td>
			<td class="listRowHead" align=right><b>Perc.</b></td>
			<td class="listRowHead" align=right><b>Amt.</b></td>
			<td class="listRowHead" align=center><b>Stat</b></td>
			<td class="listRowHead" align=center><b>Once</b></td>
			<td class="listRowHead" nowrap><b>Date (dd/mm/yy)</b></td>
			<td class="listRowHead">&nbsp;</td>
			<td class="listRowHead" nowrap align=center>
				<input type="checkbox" name="checkAll" value="1" onclick="javascript:CheckAll(this.form);">
			</td>
		</tr>
<%
		rowColor = col2
		do while not rs.eof and count < rs.pageSize
%>
			<tr>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("discCode")%></td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap><%=moneyD(rs("discFromAmt"))%></td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap><%=moneyD(rs("discToAmt"))%></td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap>
<%
					if isNull(rs("discPerc")) then
						Response.Write "-"
					else
						Response.Write formatNumber(rs("discPerc"),2) & "%"
					end if
%>
				</td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap>
<%
					if isNull(rs("discAmt")) then
						Response.Write "-"
					else
						Response.Write moneyD(rs("discAmt"))
					end if
%>
				</td>
				<td bgcolor="<%=rowColor%>" valign=top align=center><%=UCase(rs("discStatus"))%></td>
				<td bgcolor="<%=rowColor%>" valign=top align=center><%=UCase(rs("discOnceOnly"))%></td>
				<td bgcolor="<%=rowColor%>" valign=top><%=formatIntDate(rs("discValidFrom")) & " - " & formatIntDate(rs("discValidTo"))%></td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap>
					[ 
					<a href="SA_disc_edit.asp?action=edit&recid=<%=rs("idDiscOrder")%>">edit</a> 
					]
				</td>
				<td align=middle valign=top bgcolor="<%=rowColor%>">
					<input type=checkbox name="idDiscOrder" id="idDiscOrder" value="<%=rs("idDiscOrder")%>">
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
			<td nowrap colspan=5 class="listRowBot">
<%
				call pageNavigation("selectPageBot")
%>
			</td>
			<td nowrap colspan=5 align=right class="listRowBot">
				<input type=hidden name="action" id="action" value="bulkDel">
				Delete Selected Discounts? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
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

	<b>Overview</b> - There are two disctinct type of discounts, 
	namely Order discounts, and Product (item) discounts. Order 
	discounts are applied to the entire order, while Product 
	discounts are selectively applied to individual items in an order. 
	This function only applies to Order discounts, while Product 
	discounts are configured using the 
	<a href="SA_prod.asp?resetCookie=1">Product Maintenance</a> 
	function.<br><br>
	
	Each order discount is assigned a unique 
	code that must be entered in order to qualify for the discount. 
	In addition, order discounts can be given a date range and 
	can be configured so that they're only used once.<br><br>
	
	Order discounts are calculated based upon the total value of an 
	order, BEFORE taxes and shipping and AFTER product discounts 
	have been applied (if any). You specify what order total 
	will qualify for the order discount by entering a "From" and "To" 
	amount. You also have to enter what percentage or amount of the 
	order total to discount. If the customer then creates an order, 
	and enter a valid discount code, the system will check the order 
	total to see if it qualifies for the selected discount code. 
	If it does, the appropriate amount will be deducted from the 
	order total. In the list of discounts above, the following columns 
	are displayed :<br><br>
	
	<b>Code - </b>Discount code that must be entered by customer.<br>
	<b>From - </b>Minimum order amount to qualify for discount.<br>
	<b>To - </b>Maximum order amount to qualify for discount.<br>
	<b>Perc. - </b>Percentage that will be subtracted from order.<br>
	<b>Amt. - </b>Amount that will be subtracted from order.<br>
	<b>Stat - </b>Discount Status. (A)ctive, (I)inactive, (U)sed.<br>
	<b>Once - </b>Use once only? (Y)es, (N)o. If set to Yes, the discount 
	can only be used once.<br>
	<b>Date Valid - </b>Date range for which discount is valid.<br><br>
	
	<b>Add Discount</b> - Click on the "Add Discount" button, and 
	complete the form as indicated to add a new discount record.<br><br>
	
	<b>Find Discount(s)</b> - You can limit the list of Discounts  
	on this page by selecting one of several discount options, or 
	combination of discount options by using the "Find" functions.<br><br>
	
	<b>Edit Discount</b> - Click on "edit" to change discount 
	information.<br><br>
	
	<b>Delete Discount</b> - Check the box next to the discount you 
	want to delete and click the button at the bottom of the page.<br><br>
	
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
	            & "&showOnceOnly="	& server.URLEncode(showOnceOnly) _
	            & "&showPhrase="	& server.URLEncode(showPhrase)
end function
'*********************************************************************
'Make Cookie Value for Paging
'*********************************************************************
function navCookie(pageNum)

	navCookie = pageNum		  & "*|*" _
	          & showStatus	  & "*|*" _
	          & showOnceOnly  & "*|*" _
	          & showPhrase
end function
'*********************************************************************
'Format the internal integer date
'*********************************************************************
function formatIntDate(str1)
	
	if len(trim(str1)) >= 8 and isnumeric(str1) then
		formatIntDate = "" _
			& mid(str1,7,2) & "/" _
			& mid(str1,5,2) & "/" _
			& mid(str1,1,4)
	else
		formatIntDate = str1
	end if

end function
'*********************************************************************
'Display page navigation
'*********************************************************************
sub pageNavigation(formFieldName)
	Response.Write "Page "
	Response.Write "<select onChange=""location.href=this.options[selectedIndex].value"" name=" & trim(formFieldName) & ">"
	for I = 1 to TotalPages
		Response.Write "<option value=""SA_disc.asp" & navQueryStr(I) & "&sortField=" & server.URLEncode(sortField) & """ " & checkMatch(curPage,I) & ">" & I & "</option>" & vbCrlf
	next
	Response.Write "</select>&nbsp;of&nbsp;" & TotalPages & "&nbsp;&nbsp;"
	Response.Write "[&nbsp;"
	if curPage > 1 then
		Response.Write "<a href=""SA_disc.asp" & navQueryStr(curPage-1) & "&sortField=" & server.URLEncode(sortField) & """>Back</a>"
	else
		Response.Write "Back"
	end if
	Response.Write "&nbsp;|&nbsp;"
	if curPage < TotalPages then
		Response.Write "<a href=""SA_disc.asp" & navQueryStr(curPage+1) & "&sortField=" & server.URLEncode(sortField) & """>Next</a>"
	else
		Response.Write "Next"
	end if
	Response.Write "&nbsp;]"
end sub
%>