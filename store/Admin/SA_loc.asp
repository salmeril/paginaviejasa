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
dim mySQL, cn, rs, rs2

'Locations
dim idLocation
dim locName
dim locCountry
dim locState
dim locTax
dim locShipZone
dim locStatus

'Work Fields
dim I
dim item
dim count
dim pageSize
dim totalPages
dim showArr
dim sortField

dim curPage
dim showField
dim showStart
dim showTax
dim showZone

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

'Get showField
showField = Request.Form("showField")					'Form
if len(showField) = 0 then
	showField = Request.QueryString("showField")		'QueryString
end if

'Get showStart
showStart = Request.Form("showStart")					'Form
if len(showStart) = 0 then
	showStart = Request.QueryString("showStart")		'QueryString
end if

'Get showTax
showTax = Request.Form("showTax")						'Form
if len(showTax) = 0 then
	showTax = Request.QueryString("showTax")			'QueryString
end if

'Get showZone
showZone = Request.Form("showZone")						'Form
if len(showZone) = 0 then
	showZone = Request.QueryString("showZone")			'QueryString
end if

'Check if a Cookie Reset was requested
if Request.QueryString("resetCookie") = "1" then
	Response.Cookies("LocSearch").expires = Date() - 30
else
	'Check if a Cookie Recall was requested
	if Request.QueryString("recallCookie") = "1" then
		for each item in Request.Cookies
			if item = "LocSearch" then
				showArr		= Split(Request.Cookies(item),"*|*")
				curPage		= showArr(0)
				showField	= showArr(1)
				showStart	= showArr(2)
				showTax		= showArr(3)
				showZone	= showArr(4)
			end if
		next
	else
		'Save Search Criteria in a Cookie
		Response.Cookies("LocSearch") = navCookie(curPage)
		Response.Cookies("LocSearch").expires = Date() + 30
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
if len(showField) = 0 then
	showField = "locName"
end if

'Check what we will be sorting the results on
sortField = Request.Form("sortField")					'Form
if len(sortField) = 0 then
	sortField = Request.QueryString("sortField")		'QueryString
end if
if len(sortField) = 0 then
	sortField = "locName"
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Location & Tax Maintenance</font></b>
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
			<a href="SA_loc_edit.asp?action=add">Add New Country</a>
		</td>
	
		<form method="post" action="SA_loc.asp" name="form3">
		<td align=right valign=top nowrap>
			Show Countries where&nbsp;
			<select name=showField id=showField size=1>
				<option value="locName"    <%=checkMatch(showField,"locName")   %>>Country Name</option>
				<option value="locCountry" <%=checkMatch(showField,"locCountry")%>>Country Code</option>
			</select>&nbsp;
			begins with&nbsp;
			<select name=showStart id=showStart size=1>
				<option value=""  <%=checkMatch(showStart,"") %>>All</option>
				<option value="A" <%=checkMatch(showStart,"A")%>>A</option>
				<option value="B" <%=checkMatch(showStart,"B")%>>B</option>
				<option value="C" <%=checkMatch(showStart,"C")%>>C</option>
				<option value="D" <%=checkMatch(showStart,"D")%>>D</option>
				<option value="E" <%=checkMatch(showStart,"E")%>>E</option>
				<option value="F" <%=checkMatch(showStart,"F")%>>F</option>
				<option value="G" <%=checkMatch(showStart,"G")%>>G</option>
				<option value="H" <%=checkMatch(showStart,"H")%>>H</option>
				<option value="I" <%=checkMatch(showStart,"I")%>>I</option>
				<option value="J" <%=checkMatch(showStart,"J")%>>J</option>
				<option value="K" <%=checkMatch(showStart,"K")%>>K</option>
				<option value="L" <%=checkMatch(showStart,"L")%>>L</option>
				<option value="M" <%=checkMatch(showStart,"M")%>>M</option>
				<option value="N" <%=checkMatch(showStart,"N")%>>N</option>
				<option value="O" <%=checkMatch(showStart,"O")%>>O</option>
				<option value="P" <%=checkMatch(showStart,"P")%>>P</option>
				<option value="Q" <%=checkMatch(showStart,"Q")%>>Q</option>
				<option value="R" <%=checkMatch(showStart,"R")%>>R</option>
				<option value="S" <%=checkMatch(showStart,"S")%>>S</option>
				<option value="T" <%=checkMatch(showStart,"T")%>>T</option>
				<option value="U" <%=checkMatch(showStart,"U")%>>U</option>
				<option value="V" <%=checkMatch(showStart,"V")%>>V</option>
				<option value="W" <%=checkMatch(showStart,"W")%>>W</option>
				<option value="X" <%=checkMatch(showStart,"X")%>>X</option>
				<option value="Y" <%=checkMatch(showStart,"Y")%>>Y</option>
				<option value="Z" <%=checkMatch(showStart,"Z")%>>Z</option>
				<option value="0" <%=checkMatch(showStart,"0")%>>0</option>
				<option value="1" <%=checkMatch(showStart,"1")%>>1</option>
				<option value="2" <%=checkMatch(showStart,"2")%>>2</option>
				<option value="3" <%=checkMatch(showStart,"3")%>>3</option>
				<option value="4" <%=checkMatch(showStart,"4")%>>4</option>
				<option value="5" <%=checkMatch(showStart,"5")%>>5</option>
				<option value="6" <%=checkMatch(showStart,"6")%>>6</option>
				<option value="7" <%=checkMatch(showStart,"7")%>>7</option>
				<option value="8" <%=checkMatch(showStart,"8")%>>8</option>
				<option value="9" <%=checkMatch(showStart,"9")%>>9</option>
			</select>&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
		
	</tr>
	
	<tr>
		
		<form method="post" action="SA_loc.asp" name="form4">
		<td align=right valign=top nowrap colspan=2>
			Show Countries where Tax&nbsp;
			<select name=showTax id=showTax size=1>
				<option value=""  <%=checkMatch(showTax,"") %>>All</option>
				<option value="Y" <%=checkMatch(showTax,"Y")%>>> 0</option>
				<option value="N" <%=checkMatch(showTax,"N")%>>= 0</option>
			</select>&nbsp;
			and Shipping Zone =&nbsp;
			<select name=showZone id=showZone size=1>
				<option value="" <%=checkMatch(showZone,"")%>>All</option>
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
								 & checkMatch(showZone,rs("locShipZone")) _
								 & ">" _
								 & rs("locShipZone") _
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
	mySQL="SELECT idLocation,locCountry,locName,locTax,"_
		& "       locShipZone,locStatus " _
	    & "FROM   Locations " _
	    & "WHERE (locState IS NULL OR locState = '') " 'Countries Only
	    
	'Field Start
	if len(showStart) > 0 then
		mySQL = mySQL & "AND " & showField & " LIKE '" & showStart & "%' "
	end if
	    
	'Tax - Yes
	if showTax = "Y" then
		mySQL = mySQL & "AND locTax > 0 "
	end if

	'Tax - No
	if showTax = "N" then
		mySQL = mySQL & "AND locTax = 0 "
	end if
	
	'Zone
	if len(showZone) > 0 then
		mySQL = mySQL & "AND locShipZone = " & showZone & " "
	end if
	
	'Sort Order
	mySQL = mySQL & "ORDER BY " & sortField
	    
	set rs = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pageSize)
	if rs.eof then
%>
		<tr>
			<td align=center valign=middle>
				<br>
				<b>No Locations matched search criteria.</b>
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
			<td nowrap colspan=2 class="listRowTop">
<%
				call pageNavigation("selectPageTop")
%>
			</td>
			<td colspan=6 align=right class="listRowTop">
				Sort : 
				<select name=sortField id=sortField size=1 onChange="location.href='SA_loc.asp?recallCookie=1&sortField='+this.options[selectedIndex].value">
					<option value="locName"     <%=checkMatch(sortField,"locName")    %>>Country Name</option>
					<option value="locCountry"  <%=checkMatch(sortField,"locCountry") %>>Country Code</option>
					<option value="locShipZone" <%=checkMatch(sortField,"locShipZone")%>>Shipping Zone</option>
				</select>
			</td>
		</tr>
<%
		rowColor = col1
%>
		<form method="post" action="SA_loc_exec.asp" name="form2" id="form2">
		<tr>
			<td class="listRowHead" nowrap><b>Code</b></td>
			<td class="listRowHead" nowrap><b>Country</b></td>
			<td class="listRowHead" nowrap align=right><b>Tax&nbsp;%</b></td>
			<td class="listRowHead" nowrap><b>Zone</b></td>
			<td class="listRowHead" nowrap><b>Status</b></td>
			<td class="listRowHead" nowrap><b>States/Provinces</b></td>
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
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("locCountry")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("locName")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" align=right valign=top><%=formatnumber(rs("locTax"),2)%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("locShipZone")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("locStatus")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top>
<%
				mySQL = "SELECT   locState, locName " _
				      & "FROM     locations " _
				      & "WHERE    locCountry = '" & rs("locCountry") & "' " _
				      & "AND      NOT(locState IS NULL OR locState = '') " _
				      & "ORDER BY locName "
				set rs2 = openRSexecute(mySQL)
				if not rs2.eof then
					Response.Write "<select name=locState" & rs("locCountry") & " size=1>"
					do while not rs2.EOF
						Response.Write "<option value=""" & rs2("locState") & """>" & rs2("locName") & "</option>" & vbCrLf
						rs2.MoveNext
					loop
					Response.Write "</select>"
				else
					Response.Write "None"
				end if
				call closeRS(rs2)
%>
				</td>
				<td bgcolor="<%=rowColor%>" align=right valign=top nowrap>
					[ 
					<a href="SA_loc_edit.asp?action=edit&recid=<%=rs("idLocation")%>">edit</a> 
					]
				</td>
				<td align=middle valign=top bgcolor="<%=rowColor%>">
					<input type=checkbox name="locCountry" id="locCountry" value="<%=rs("locCountry")%>">
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
			<td nowrap colspan=2 class="listRowBot">
<%
				call pageNavigation("selectPageBot")
%>
			</td>
			<td colspan=6 align=right class="listRowBot">
				<input type=hidden name="action" id="action" value="bulkDel">
				Delete Selected Locations? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
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

	<b>Add Country</b> - Click on the Add New Country button, and 
	complete the form as indicated. Once you have added the Country, you 
	will be able to Add States or Provinces for that Country.<br><br>
	<b>Find Country(s)</b> - You have several options for finding a specific 
	Country.<br><br>
	
	1. You can list all Countries where the Country Name or Country Code 
	fields start with the specified Alphabetic or Numeric value.<br><br>
	
	2. You can list al Countries with the selected Tax and/or Shipping Zone 
	assignments.<br><br>
	
	With all the searches above, you may have to page through the list 
	of Countries if a sufficient number of Countries satisified the search 
	criteria you specified.<br><br>
	
	<b>Edit Country</b> - Click to change Country's information. You 
	can also Add, Remove and Edit States/Provinces for that Country from 
	the Country Edit page.<br><br>
	
	<b>Delete Country</b> - Check the box next to the countries you want 
	to delete, and click "Yes". When you Delete a country, 
	any related State and Province records will also be deleted.<br><br>
	
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

	navQueryStr = "?curPage="	& server.URLEncode(pageNum) _
	            & "&showField="	& server.URLEncode(showField) _
	            & "&showStart="	& server.URLEncode(showStart) _
	            & "&showTax="	& server.URLEncode(showTax) _
	            & "&showZone="	& server.URLEncode(showZone)

end function
'*********************************************************************
'Make Cookie Value for Paging
'*********************************************************************
function navCookie(pageNum)

	navCookie = pageNum		& "*|*" _
	          & showField	& "*|*" _
	          & showStart	& "*|*" _
	          & showTax		& "*|*" _
	          & showZone
	          
end function
'*********************************************************************
'Display page navigation
'*********************************************************************
sub pageNavigation(formFieldName)
	Response.Write "Page "
	Response.Write "<select onChange=""location.href=this.options[selectedIndex].value"" name=" & trim(formFieldName) & ">"
	for I = 1 to TotalPages
		Response.Write "<option value=""SA_loc.asp" & navQueryStr(I) & "&sortField=" & server.URLEncode(sortField) & """ " & checkMatch(curPage,I) & ">" & I & "</option>" & vbCrlf
	next
	Response.Write "</select>&nbsp;of&nbsp;" & TotalPages & "&nbsp;&nbsp;"
	Response.Write "[&nbsp;"
	if curPage > 1 then
		Response.Write "<a href=""SA_loc.asp" & navQueryStr(curPage-1) & "&sortField=" & server.URLEncode(sortField) & """>Back</a>"
	else
		Response.Write "Back"
	end if
	Response.Write "&nbsp;|&nbsp;"
	if curPage < TotalPages then
		Response.Write "<a href=""SA_loc.asp" & navQueryStr(curPage+1) & "&sortField=" & server.URLEncode(sortField) & """>Next</a>"
	else
		Response.Write "Next"
	end if
	Response.Write "&nbsp;]"
end sub
%>