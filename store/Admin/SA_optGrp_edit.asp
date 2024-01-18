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
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="../Scripts/_INCappFunctions_.asp"-->
<!--#include file="_INCoptions_.asp"-->
<%
'Database
dim mySQL, cn, rs

'OptionsGroups
dim idOptionGroup
dim optionGroupDesc
dim optionReq
dim optionType

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
	<b><font size=3>Option Group Maintenance</font></b>
	<br><br>
</P>

<%
'Page Tabs
call optTabs("OG")

'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if action <> "edit" and action <> "del" and action <> "add" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idOptionGroup
if action = "edit" or action = "del" then
	idOptionGroup = trim(Request.QueryString("recId"))
	if len(idOptionGroup) = 0 then
		idOptionGroup = trim(Request.Form("recId"))
	end if
	if idOptionGroup = "" or not isNumeric(idOptionGroup) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if
end if

'Get Record
if action = "edit" or action = "del" then
	mySQL="SELECT * " _
	    & "FROM   OptionsGroups " _
	    & "WHERE  idOptionGroup = " & idOptionGroup
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	else
		OptionGroupDesc = trim(rs("OptionGroupDesc"))
		OptionReq       = trim(rs("OptionReq"))
		OptionType      = trim(rs("OptionType"))
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
	<span class="textBlockHead">Edit Option Group</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td><b>Group Description</b></td>
			<td><b>Type</b></td>
			<td><b>Required?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_optGrp_exec.asp" name="form1">
			<td>
				<input type=text name=optionGroupDesc id=optionGroupDesc size=30 maxlength=50 value="<%=optionGroupDesc%>">
			</td>
			<td>
				<select name=optionType id=optionType size=1>
					<option value="S" <%=checkMatch(optionType,"S")%>>Drop-down List</option>
					<option value="T" <%=checkMatch(optionType,"T")%>>Text Input</option>
				</select>
			</td>
			<td>
				<select name=optionReq id=optionReq size=1>
					<option value="N" <%=checkMatch(optionReq,"N")%>>No</option>
					<option value="Y" <%=checkMatch(optionReq,"Y")%>>Yes</option>
				</select>
			</td>
			<td>
				<input type=hidden name=idOptionGroup id=idOptionGroup value="<%=idOptionGroup%>">
				<input type=hidden name=action        id=action        value="edit">
				<input type=submit name=submit1       id=submit1       value="Update">
			</td>
			</form>
		</tr>
		<tr>
			<td colspan=4>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=4 bgcolor="#dddddd">
				<span class="textBlockHead">Options</span>
			</td>
		</tr>
		<tr>
			<td colspan=4>
				<table border=0 cellspacing=0 cellpadding=5 bgcolor=#eeeeee>
<%
					mySQL="SELECT a.idOptOptGroup, b.idOption, b.optionDescrip " _
						& "FROM   optionsXref a " _
						& "INNER JOIN options b " _
						& "ON     a.idoption = b.idoption " _
						& "WHERE  a.idOptionGroup = " & idOptionGroup
					set rs = openRSexecute(mySQL)
					do while not rs.eof
%>
						<tr>
							<td nowrap><%=rs("optionDescrip")%></td>
							<td nowrap>
								<a href="SA_optGrp_exec.asp?action=delOpt&idOptionGroup=<%=idOptionGroup%>&recId=<%=rs("idOptOptGroup")%>">Remove</a>
							</td>
						</tr>
<%
						rs.movenext
					loop
					call closeRS(rs)
%>
					<tr>
						<form method="post" action="SA_optGrp_exec.asp" name="form3">
						<td>
							<select name=idOption id=idOption size=1>
								<option value="">-- Select --</option>
<%
								'Get available Options for Option Group
								mySQL="SELECT a.idOption, a.optionDescrip " _
								    & "FROM   options a " _
								    & "WHERE  NOT EXISTS " _
								    & "      (SELECT b.idOptOptGroup " _
								    & "       FROM   optionsXref b " _
								    & "       WHERE  b.idoption = a.idOption " _
								    & "       AND    b.idOptionGroup = " & idOptionGroup & ")" _
								    & "ORDER BY a.optionDescrip"
								set rs = openRSexecute(mySQL)
								do while not rs.eof
									Response.Write "<option value=""" & rs("idOption") & """>" & rs("optionDescrip") & "</option>"
									rs.movenext
								loop
								call closeRS(rs)
%>
							</select>
						</td>
						<td>
							<input type=hidden name=idOptionGroup id=idOptionGroup value="<%=idOptionGroup%>">
							<input type=hidden name=action        id=action        value="addOpt">
							<input type=submit name=submit1       id=submit1       value=" Add ">
						</td>
						</form>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<%
end if

'Add
if action = "add" then
%>
	<span class="textBlockHead">Add Option Group</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<form method="post" action="SA_optGrp_exec.asp" name="form1">
		<tr>
			<td><b>Group Description</b></td>
			<td><b>Type</b></td>
			<td><b>Required?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>
				<input type=text name=optionGroupDesc id=optionGroupDesc size=30 maxlength=50>
			</td>
			<td>
				<select name=optionType id=optionType size=1>
					<option value="S">Drop-down List</option>
					<option value="T">Text Input</option>
				</select>
			</td>
			<td>
				<select name=optionReq id=optionReq size=1>
					<option value="N">No</option>
					<option value="Y">Yes</option>
				</select>
			</td>
			<td>
				<input type=hidden name=action  id=action  value="add">
				<input type=submit name=submit1 id=submit1 value="Add">
			</td>
		</tr>
	</form>
	</table>
<%
end if

'Delete
if action = "del" then
%>
	<span class="textBlockHead">Delete Option Group</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<form method="post" action="SA_optGrp_exec.asp" name="form2">
		<tr>
			<td><b>Group Description</b></td>
			<td><b>Type</b></td>
			<td><b>Required?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td><%=optionGroupDesc%></td>
			<td><%=optionType%></td>
			<td><%=optionReq%></td>
			<td>
				<input type=hidden name=idOptionGroup id=idOptionGroup value="<%=idOptionGroup%>">
				<input type=hidden name=action        id=action        value="del">
				<input type=submit name=submit1       id=submit1       value="Delete">
			</td>
		</tr>
	</form>
	</table>
<%
end if

if action = "edit" or action = "add" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>
	
		<b>Option Group Description</b> - Mandatory. A name for the 
		option group.<br><br>
		
		<b>Type</b> - Mandatory. Drop-down lists will be displayed 
		as a drop-down box in which multiple options are displayed 
		and one is selected. Text input will display a text box into 
		which the customer must enter some text that goes with the 
		option (eg. an inscription on a bracelet, etc.). This type 
		of option group is limited to one option per group.<br><br>
		
		<b>Required?</b> - Mandatory. If set to 'Yes', the customer 
		will be forced to pick at least one of the options in the 
		option group when they order an item. If set to 'No', the 
		options for this option group will be optional.<br><br>
<%
		if action = "edit" then
%>
		<b>Options</b> - This is a list of options currently linked 
		to this option group. You can add more options to the option 
		group, or remove them from the option group.<br><br>
		
		Note : When you add the option group to a product, you 
		will also be given the opportunity to exclude any of the 
		options in this option group for that particular product.
		<br><br>
<%
		end if
%>
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
	<a href=SA_optGrp.asp>List Groups</a> | 
	<a href=SA_optGrp_edit.asp?action=edit&recID=<%=idOptionGroup%>>Edit</a> | 
	<a href=SA_optGrp_edit.asp?action=add>Add</a> | 
	<a href=SA_optGrp_edit.asp?action=del&recID=<%=idOptionGroup%>>Delete</a> 
	]
<%
end sub
%>
