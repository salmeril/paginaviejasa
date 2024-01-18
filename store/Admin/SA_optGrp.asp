<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Options Groups Maintenance
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
<!--#include file="_INCoptions_.asp"-->
<%
'Database
dim mySQL, cn, rs

'OptionsGroups
dim idOptionGroup
dim optionGroupDesc
dim optionReq
dim optionType

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
	<b><font size=3>Option Group Maintenance</font></b>
	<br><br>
</P>

<%
'Page Tabs
call optTabs("OG")

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
			<a href="SA_optGrp_edit.asp?action=add">Add New Option Group</a>
		</td>
	</tr>
</table>

<br>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="listTable">
	<tr>
		<td colspan=7 class="listRowTop">&nbsp;</td>
	</tr>
<%
	rowColor = col1
%>
	<tr>
		<td class="listRowHead"><b>ID</b></td>
		<td class="listRowHead"><b>Description</b></td>
		<td class="listRowHead"><b>Type</b></td>
		<td class="listRowHead"><b>Req</b></td>
		<td class="listRowHead"><b>Opt</b></td>
		<td class="listRowHead"><b>Prod</b></td>
		<td class="listRowHead" width="1%"><b>&nbsp;</b></td>
	</tr>
<%
	rowColor = col2

	'Retrieve all Option Groups
	mySQL="SELECT a.*, " _
		& "      (SELECT COUNT(*) " _
		& "       FROM   optionsXref b " _
		& "       WHERE  b.idOptionGroup = a.idOptionGroup) " _
		& "       AS OptionCount, " _ 
		& "      (SELECT COUNT(*) " _
		& "       FROM   optionsGroupsXref c " _
		& "       WHERE  c.idOptionGroup = a.idOptionGroup) " _
		& "       AS ProductCount " _ 
	    & "FROM   OptionsGroups a " _
	    & "ORDER BY a.optionGroupDesc"
	set rs = openRSexecute(mySQL)
	do while not rs.eof
%>
		<tr>
			<td bgcolor="<%=rowColor%>"><%=rs("idOptionGroup")%></td>
			<td bgcolor="<%=rowColor%>"><%=rs("optionGroupDesc")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>"><%=rs("optionType")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>"><%=rs("optionReq")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>"><%=rs("optionCount")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>"><%=rs("productCount")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>" nowrap>
				[ 
				<a href="SA_optGrp_edit.asp?action=edit&recid=<%=rs("idOptionGroup")%>">edit</a> | 
<%
				if rs("OptionCount") > 0 or rs("ProductCount") > 0 then
%>
				delete 
<%
				else
%>
				<a href="SA_optGrp_edit.asp?action=del&recid=<%=rs("idOptionGroup")%>">delete</a> 
<%
				end if
%>
				]
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
		<td colspan=7 class="listRowBot">&nbsp;</td>
	</tr>
<%
	call closeRS(rs)
%>
</table>

<br>

<span class="textBlockHead">Help and Instructions :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr><td>

	<b>Overview</b> - An Option Group is a collection of one or more 
	options, grouped together under a common name. Once an option group 
	has been created, it is then linked to one or more Products using 
	the <a href="SA_prod.asp?resetCookie=1">Product Maintenance</a> 
	function.<br><br>
	
	Option Groups can be displayed as follows :<br>
	<ol>
		<li>Drop-down select box containing one more options.</li>
		<li>Text box into which the customer enters additional text.</li>
	</ol>
	
	In the example below, the option groups "Color" and "Size" will be 
	displayed as drop-down select boxes allowing the the customer to 
	select a "Color" and/or "Size" from a list. The "Printed Text" 
	option group is a text box into which the customer can enter the 
	text they want to have printed on the T-Shirt.<br><br>
	
	An option group can be made mandatory. This will ensure that a 
	customer has to pick one of the options in a list, or enter some 
	text into a text box before they add a product to their shopping 
	cart.<br><br>
	
	<b>Example :</b><br>
	
	<ul>
		<li>T-Shirt
			<ul>
				<li><font color=blue>Color</font>
					<ul>
						<li>White</li>
						<li>Red</li>
						<li>Blue</li>
					</ul>
				</li>
				<li><font color=blue>Size</font>
					<ul>
						<li>S</li>
						<li>M</li>
						<li>L</li>
						<li>XL</li>
					</ul>
				</li>
				<li><font color=blue>Printed Text</font>
					<ul>
						<li>&lt;Text entered by customer&gt;</li>
					</ul>
				</li>
			</ul>
		</li>
	</ul>

	NOTE : You can not delete option groups that have options or 
	products linked to them. You will need to delete the options 
	or products first (or un-link them from the option group). 
	You can see the number of options and products linked to an 
	option group under the heading "Opt" and "Prod".<br><br>

</td></tr>
</table>
<%
call closedb()
%>
<!--#include file="_INCfooter_.asp"-->