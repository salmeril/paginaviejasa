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
dim mySQL, cn, rs, rs2

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
dim item
dim count
dim pageSize
dim totalPages
dim showArr
dim sortField

dim curPage
dim showStatus
dim showRating
dim showPhrase
dim showProd

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

'Get showRating
showRating = Request.Form("showRating")					'Form
if len(showRating) = 0 then
	showRating = Request.QueryString("showRating")		'QueryString
end if

'Get showPhrase
showPhrase = Request.Form("showPhrase")					'Form
if len(showPhrase) = 0 then
	showPhrase = Request.QueryString("showPhrase")		'QueryString
end if

'Get showProd
showProd = Request.Form("showProd")						'Form
if len(showProd) = 0 then
	showProd = Request.QueryString("showProd")			'QueryString
end if

'Check if a Cookie Reset was requested
if Request.QueryString("resetCookie") = "1" then
	Response.Cookies("ReviewSearch").expires = Date() - 30
else
	'Check if a Cookie Recall was requested
	if Request.QueryString("recallCookie") = "1" then
		for each item in Request.Cookies
			if item = "ReviewSearch" then
				showArr		= Split(Request.Cookies(item),"*|*")
				curPage		= showArr(0)
				showStatus	= showArr(1)
				showRating	= showArr(2)
				showPhrase  = showArr(3)
				showProd    = showArr(4)
			end if
		next
	else
		'Save Search Criteria in a Cookie
		Response.Cookies("ReviewSearch") = navCookie(curPage)
		Response.Cookies("ReviewSearch").expires = Date() + 30
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
	sortField = "revDate DESC"
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Review Maintenance</font></b>
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
		<form method="post" action="SA_rev.asp" name="form4">
		<td align=right valign=top nowrap>
			Status = 
			<select name=showStatus id=showStatus size=1>
				<option value=""  <%=checkMatch(showStatus,"") %>>All</option>
				<option value="A" <%=checkMatch(showStatus,"A")%>>Active</option>
				<option value="I" <%=checkMatch(showStatus,"I")%>>Pending</option>
				<option value="R" <%=checkMatch(showStatus,"R")%>>Rejected</option>
			</select>&nbsp;
			Rating = 
			<select name=showRating id=showRating size=1>
				<option value=""  <%=checkMatch(showRating,"") %>>All</option>
				<option value="1" <%=checkMatch(showRating,"1")%>>1</option>
				<option value="2" <%=checkMatch(showRating,"2")%>>2</option>
				<option value="3" <%=checkMatch(showRating,"3")%>>3</option>
				<option value="4" <%=checkMatch(showRating,"4")%>>4</option>
				<option value="5" <%=checkMatch(showRating,"5")%>>5</option>
			</select>&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
	</tr>
	<tr>
		<form method="post" action="SA_rev.asp" name="form2">
		<td align=right valign=top nowrap>
			Show Reviews containing the phrase 
			<input type=text name=showPhrase id=showPhrase size=20 maxlength=50 value="<%=showPhrase%>">&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
	</tr>
	<tr>
		<form method="post" action="SA_rev.asp" name="form1">
		<td align=right valign=top nowrap>
			Show all reviews for 
			<select name=showProd id=showProd size=1>
				<option value="" <%=checkMatch(showProd,"")%>>-- Select Product with 1 or more reviews --</option>
<%
				mySQL = "SELECT   a.idProduct, a.SKU, a.description " _
				      & "FROM     Products a " _
				      & "WHERE    EXISTS (SELECT b.idProduct FROM reviews b WHERE b.idProduct = a.idProduct) " _
				      & "ORDER BY a.description "
				set rs = openRSexecute(mySQL)
				do while not rs.EOF
					Response.Write "<option value=""" _
								 & rs("idProduct") _
								 & """ " _
								 & checkMatch(showProd,rs("idProduct")) _
								 & ">" _
								 & rs("description") _
								 & "</option>"
					rs.MoveNext
				loop
				call closeRS(rs)
%>
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
	mySQL="SELECT a.idReview,a.revDate,a.revAuditInfo,a.revStatus," _
		& "       a.revRating,a.revName,a.revLocation,a.revSubj," _
		& "       b.description " _
	    & "FROM   reviews a, products b " _
	    & "WHERE  a.idProduct = b.idProduct "
	    
	'revStatus
	if len(showStatus) > 0 then
		mySQL = mySQL & "AND revStatus = '" & showStatus & "' "
	end if
	    
	'revRating
	if len(showRating) > 0 then
		mySQL = mySQL & "AND revRating = " & showRating & " "
	end if
	
	'Search Phrase
	if len(showPhrase) > 0 then	
		mySQL = mySQL _
			& "AND (revSubj      LIKE '%" & replace(showPhrase,"'","''") & "%'  " _
			& "OR   revDetail    LIKE '%" & replace(showPhrase,"'","''") & "%'  " _
			& "OR   revName      LIKE '%" & replace(showPhrase,"'","''") & "%'  " _
			& "OR   revLocation  LIKE '%" & replace(showPhrase,"'","''") & "%'  " _
			& "OR   revAuditInfo LIKE '%" & replace(showPhrase,"'","''") & "%') "
	end if
	
	'idProduct
	if len(showProd) > 0 then
		mySQL = mySQL & "AND a.idProduct = " & showProd & " "
	end if
	
	'Sort Order
	mySQL = mySQL & "ORDER BY " & sortField
	
	set rs = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pageSize)
	if rs.eof then
%>
		<tr>
			<td align=center valign=middle>
				<br>
				<b>No Reviews matched search criteria.</b>
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
			<td colspan=1 class="listRowTop">
<%
				call pageNavigation("selectPageTop")
%>
			</td>
			<td colspan=4 align=right class="listRowTop">
				Sort : 
				<select name=sortField id=sortField size=1 onChange="location.href='SA_rev.asp?recallCookie=1&sortField='+this.options[selectedIndex].value">
					<option value="revDate DESC" <%=checkMatch(sortField,"revDate DESC")%>>Date (Descending)</option>
					<option value="revDate"      <%=checkMatch(sortField,"revDate")     %>>Date (Ascending)</option>
					<option value="revName"      <%=checkMatch(sortField,"revName")     %>>Reviewer Name</option>
					<option value="revRating"    <%=checkMatch(sortField,"revRating")   %>>Rating</option>
					<option value="revStatus"    <%=checkMatch(sortField,"revStatus")   %>>Status</option>
				</select>
			</td>
		</tr>
<%
		rowColor = col1
%>
		<form method="post" action="SA_rev_exec.asp" name="form3" id="form3">
		<tr>
			<td class="listRowHead"><b>Review Summary</b></td>
			<td class="listRowHead"><b>Status</b></td>
			<td class="listRowHead"><b>Rating</b></td>
			<td class="listRowHead">&nbsp;</td>
			<td class="listRowHead" align=center>
				<input type="checkbox" name="checkAll" value="1" onclick="javascript:CheckAll(this.form);">
			</td>
		</tr>
<%
		rowColor = col2
		do while not rs.eof and count < rs.pageSize
%>
			<tr>
				<td bgcolor="<%=rowColor%>" valign=top>
					<b><%=formatTheDate(rs("revDate"))%> - <%=server.HTMLEncode(rs("revSubj") & "")%></b><br>
					<span style="color: #800000;"><%=server.HTMLEncode(rs("revName") & "") & " - " & server.HTMLEncode(rs("revLocation") & "")%></span><br>
					<%=rs("description")%><br>
					IP : <a href="http://www.samspade.org/t/lookat?a=<%=rs("revAuditInfo")%>" target="_blank"><%=rs("revAuditInfo")%></a><br>
				</td>
				<td bgcolor="<%=rowColor%>" valign=top>
<%
					select case trim(UCase(rs("revStatus")))
					case "A"
						Response.Write "<b><font color=green>ACT</font></b>"
					case "I"
						Response.Write "<b>PEN</b>"
					case "R"
						Response.Write "<b><font color=red>REJ</font></b>"
					case else
						Response.Write "<b>" & rs("revStatus") & "</b>"
					end select
%>
				</td>
				<td bgcolor="<%=rowColor%>" valign=top><b><%=ratingImage(rs("revRating"))%></b></td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap>
					[ 
					<a href="SA_rev_edit.asp?action=edit&recid=<%=rs("idReview")%>">edit</a> 
					]
				</td>
				<td align=middle valign=top bgcolor="<%=rowColor%>">
					<input type=checkbox name="idReview" id="idReview" value="<%=rs("idReview")%>">
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
			<td nowrap colspan=1 class="listRowBot">
<%
				call pageNavigation("selectPageBot")
%>
			</td>
			<td nowrap colspan=4 align=right class="listRowBot">
				<input type=hidden name="action" id="action" value="bulkDel">
				Delete Selected Reviews? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
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

	<b>Overview</b> - Product reviews are allowed/disallowed per 
	product using the <a href="SA_prod.asp?resetCookie=1">Product Maintenance</a> 
	function. You can also specify if new reviews will be "activated" 
	immediately after being entered by the customer, or if they need 
	to be approved first.<br><br>
	
	<b>Status</b> - The Status field will tell you if a review is 
	active, pending, or rejected.<br><br>
	
	<b><font color=green>ACT</font></b> - Active review. This 
	review will be displayed in the list of reviews for the product.<br>
	
	<b>PEN</b> - Pending (Inactive) review. This review is awaiting 
	authorization before being viewable.<br>
	
	<b><font color=red>REJ</font></b> - Rejected review. The review 
	has been deliberately removed from the list of viewable reviews.<br><br>
	
	Note : To prevent "spamming", the software checks the IP 
	address of the customer against existing reviews for a particular 
	product. It is therefore suggested that you do not delete unwanted 
	reviews, but mark them as "rejected" instead. This will ensure that 
	the customer will be unable to add another review for that product, 
	as long as they use the same IP address.<br><br>
	
	<b>Find Review(s)</b> - You can limit the list of Reviews 
	on this page by selecting one of several review options, or 
	combination of options by using the "Find" functions.<br><br>
	
	<b>Edit Review</b> - Click on "edit" to change a review.<br><br>
	
	<b>Delete Review</b> - Check the box next to any reviews you 
	want to delete, and click the button at the bottom of the list 
	of reviews.<br><br>
	
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

	navQueryStr = "?curPage="	 & server.URLEncode(pageNum) _
	            & "&showStatus=" & server.URLEncode(showStatus) _
	            & "&showRating=" & server.URLEncode(showRating) _
	            & "&showPhrase=" & server.URLEncode(showPhrase) _
	            & "&showProd="	 & server.URLEncode(showProd)
end function
'*********************************************************************
'Make Cookie Value for Paging
'*********************************************************************
function navCookie(pageNum)

	navCookie = pageNum	    & "*|*" _
	          & showStatus  & "*|*" _
	          & showRating  & "*|*" _
	          & showPhrase  & "*|*" _
	          & showProd
end function
'*********************************************************************
'Display page navigation
'*********************************************************************
sub pageNavigation(formFieldName)
	Response.Write "Page "
	Response.Write "<select onChange=""location.href=this.options[selectedIndex].value"" name=" & trim(formFieldName) & ">"
	for I = 1 to TotalPages
		Response.Write "<option value=""SA_rev.asp" & navQueryStr(I) & "&sortField=" & server.URLEncode(sortField) & """ " & checkMatch(curPage,I) & ">" & I & "</option>" & vbCrlf
	next
	Response.Write "</select>&nbsp;of&nbsp;" & TotalPages & "&nbsp;&nbsp;"
	Response.Write "[&nbsp;"
	if curPage > 1 then
		Response.Write "<a href=""SA_rev.asp" & navQueryStr(curPage-1) & "&sortField=" & server.URLEncode(sortField) & """>Back</a>"
	else
		Response.Write "Back"
	end if
	Response.Write "&nbsp;|&nbsp;"
	if curPage < TotalPages then
		Response.Write "<a href=""SA_rev.asp" & navQueryStr(curPage+1) & "&sortField=" & server.URLEncode(sortField) & """>Next</a>"
	else
		Response.Write "Next"
	end if
	Response.Write "&nbsp;]"
end sub
%>