<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Customer Maintenance
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

'Customer
dim idCust
dim status
dim dateCreated
dim dateCreatedInt
dim name
dim lastName
dim customerCompany
dim phone
dim email
dim password
dim address
dim city
dim locState
dim locCountry
dim zip
dim paymentType
dim shippingName
dim shippingLastName
dim shippingAddress
dim shippingCity
dim shippingLocState
dim shippingLocCountry
dim shippingZip
dim futureMail
dim generalComments

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
	Response.Cookies("CustSearch").expires = Date() - 30
else
	'Check if a Cookie Recall was requested
	if Request.QueryString("recallCookie") = "1" then
		for each item in Request.Cookies
			if item = "CustSearch" then
				showArr		= Split(Request.Cookies(item),"*|*")
				curPage		= showArr(0)
				showStatus	= showArr(1)
				showField	= showArr(2)
				showPhrase	= showArr(3)
			end if
		next
	else
		'Save Search Criteria in a Cookie
		Response.Cookies("CustSearch") = navCookie(curPage)
		Response.Cookies("CustSearch").expires = Date() + 30
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
	sortField = "lastName"
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Customer Maintenance</font></b>
	<br><br>
</P>

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
	
		<form method="post" action="SA_cust.asp" name="form2">
		<td align=right valign=top nowrap>
			Show Customers where&nbsp;
			<select name=showField id=showField size=1>
				<option value="">-- Select --</option>
				<option value="name"     <%=checkMatch(showField,"name")    %>>First Name</option>
				<option value="lastName" <%=checkMatch(showField,"lastName")%>>Last Name</option>
				<option value="email"    <%=checkMatch(showField,"email")   %>>Email</option>
				<option value="address"  <%=checkMatch(showField,"address") %>>Address</option>
			</select>&nbsp;
			contains the phrase&nbsp;
			<input type=text name=showPhrase id=showPhrase size=20 maxlength=50 value="<%=showPhrase%>">&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
		
	</tr>
	
	<tr>
	
		<form method="post" action="SA_cust.asp" name="form4">
		<td align=right valign=top nowrap>
			Show Customers where Status is&nbsp;
			<select name=showStatus id=showStatus size=1>
				<option value="">Show all Customers</option>
				<option value="">-----------------------</option>
				<option value="A" <%=checkMatch(showStatus,"A")%>>Active</option>
				<option value="I" <%=checkMatch(showStatus,"I")%>>InActive</option>
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
	mySQL="SELECT idCust,status,dateCreated,name,lastName,email " _
	    & "FROM   customer " _
	    & "WHERE  1=1 " 'Dummy check to set up conditional checks below
	    
	'Status
	if len(showStatus) > 0 then
		mySQL = mySQL & "AND status = '" & showStatus & "' "
	end if

	'Field and Search Phrase
	if len(showField) > 0 and len(showPhrase) > 0 then
		mySQL = mySQL & "AND " & showField & " LIKE '%" & replace(showPhrase,"'","''") & "%' "
	end if

	'Sort Order
	mySQL = mySQL & "ORDER BY " & sortField
	
	set rs = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pageSize)
	if rs.eof then
%>
		<tr>
			<td align=center valign=middle>
				<br>
				<b>No Customers matched search criteria.</b>
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
			<td colspan=3 class="listRowTop">
<%
				call pageNavigation("selectPageTop")
%>
			</td>
			<td colspan=3 align=right class="listRowTop">
				Sort : 
				<select name=sortField id=sortField size=1 onChange="location.href='SA_cust.asp?recallCookie=1&sortField='+this.options[selectedIndex].value">
					<option value="lastName"    <%=checkMatch(sortField,"lastName")   %>>Last Name</option>
					<option value="email"       <%=checkMatch(sortField,"email")      %>>Email Address</option>
					<option value="dateCreated" <%=checkMatch(sortField,"dateCreated")%>>Date Created</option>
				</select>
			</td>
		</tr>
<%
		rowColor = col1
%>
		<tr>
			<td class="listRowHead"><b>ID</b></td>
			<td class="listRowHead"><b>Name</b></td>
			<td class="listRowHead"><b>EMail</b></td>
			<td class="listRowHead"><b>Date Created</b></td>
			<td class="listRowHead"><b>Status</b></td>
			<td class="listRowHead"><b>&nbsp;</b></td>
		</tr>
<%
		rowColor = col2
		do while not rs.eof and count < rs.pageSize
%>
			<tr>
				<td bgcolor="<%=rowColor%>" valign=top nowrap><%=rs("idCust")%></td>
				<td bgcolor="<%=rowColor%>" valign=top nowrap>
					<%=rs("lastName") & ", " & rs("name")%>
				</td>
				<td bgcolor="<%=rowColor%>" valign=top nowrap><%=rs("email")%></td>
				<td bgcolor="<%=rowColor%>" valign=top nowrap><%=formatTheDate(rs("dateCreated"))%></td>
				<td bgcolor="<%=rowColor%>" valign=top nowrap><%=rs("status")%></td>
				<td bgcolor="<%=rowColor%>" align=right valign=top nowrap>
					[ 
					<a href="SA_cust_edit.asp?action=edit&recid=<%=rs("idCust")%>">edit</a> | 
					<a href="SA_cust_edit.asp?action=del&recid=<%=rs("idCust")%>">delete</a> 
					]
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
			<td nowrap colspan=6 class="listRowBot">
<%
				call pageNavigation("selectPageBot")
%>
			</td>
		</tr>
<%
	end if
	call closeRS(rs)
%>
</table>

<br>

<span class="textBlockHead">Help and Instructions :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr><td>

	<b>Find Customer(s)</b> - You have several options for finding a 
	specific Customer. Customers are display in Last Name, First Name 
	sequence.<br><br>
	
	1. You can list all Customers which contain a specific phrase in one 
	of several fields.<br><br>
	
	2. You can list all Customers with a specified Status.<br><br>
	
	<b>Edit Customer</b> - Click to change Customer's Status and/or 
	Information.<br><br>
	
	<b>Delete Customer</b> - Click to delete a Customer from the Database. 
	Please note that a Customer record can NOT be deleted if there are 
	Orders linked to it. You will have to delete the Orders linked to the 
	Customer first, then return and delete the Customer. Or you could 
	simply InActivate the Customer.<br><br>
	
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
		Response.Write "<option value=""SA_cust.asp" & navQueryStr(I) & "&sortField=" & server.URLEncode(sortField) & """ " & checkMatch(curPage,I) & ">" & I & "</option>" & vbCrlf
	next
	Response.Write "</select>&nbsp;of&nbsp;" & TotalPages & "&nbsp;&nbsp;"
	Response.Write "[&nbsp;"
	if curPage > 1 then
		Response.Write "<a href=""SA_cust.asp" & navQueryStr(curPage-1) & "&sortField=" & server.URLEncode(sortField) & """>Back</a>"
	else
		Response.Write "Back"
	end if
	Response.Write "&nbsp;|&nbsp;"
	if curPage < TotalPages then
		Response.Write "<a href=""SA_cust.asp" & navQueryStr(curPage+1) & "&sortField=" & server.URLEncode(sortField) & """>Next</a>"
	else
		Response.Write "Next"
	end if
	Response.Write "&nbsp;]"
end sub
%>