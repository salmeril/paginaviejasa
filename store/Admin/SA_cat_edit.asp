<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Category Maintenance
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

'Categories
dim idCategory
dim categoryDesc
dim idParentCategory
dim categoryFeatured
dim categoryHTML

'Work Fields
dim action
dim ParentCategoryDesc

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
	<b><font size=3>Category Maintenance</font></b>
	<br><br>
</P>

<%
'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if action <> "edit" and action <> "del" and action <> "add" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idCategory
if action = "edit" or action = "del" then
	idCategory = trim(Request.QueryString("recId"))
	if len(idCategory) = 0 then
		idCategory = trim(Request.Form("recId"))
	end if
	if idCategory = "" or not isNumeric(idCategory) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if
end if

'Get Record
if action = "edit" or action = "del" then
	mySQL="SELECT a.idCategory, a.categoryDesc, " _
		& "       a.idParentCategory, a.categoryFeatured, " _
		& "       a.categoryHTML, " _
		& "      (SELECT b.categoryDesc " _
		& "       FROM   categories b " _
		& "       WHERE  b.idCategory = a.idParentCategory) " _
		& "       AS ParentCategoryDesc " _
	    & "FROM   categories a " _
	    & "WHERE  a.idCategory = " & idCategory
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	else
		'Assign to local variables
		categoryDesc		= trim(rs("categoryDesc"))
		idParentCategory	= rs("idParentCategory")
		categoryFeatured    = rs("categoryFeatured")
		ParentCategoryDesc  = rs("ParentCategoryDesc")
		categoryHTML		= trim(rs("categoryHTML"))
		
		'CategoryFeatured was added later, so to ensure that older 
		'records have a valid value, we perform the checks below.
		if isnull(categoryFeatured) or len(categoryFeatured) = 0 then
			categoryFeatured = "N"
		end if
		if categoryFeatured <> "Y" then
			categoryFeatured = "N"
		end if
	end if
	call closeRS(rs)
end if

'Edit
if action = "edit" then
%>
	<span class="textBlockHead">Edit Category</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<form method="post" action="SA_cat_exec.asp" name="form1">
		<tr>
			<td nowrap><b>Category Description</b></td>
			<td nowrap><b>Parent Category</b></td>
			<td nowrap><b>Featured?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>
				<input type=text name=categoryDesc id=categoryDesc size=20 maxlength=50 value="<%=CategoryDesc%>">
			</td>
			<td>
<%
				if idParentCategory = 0 then
%>
				<input type=hidden name=idParentCategory id=idParentCategory value="0">
				<b>&lt;ROOT&gt;</b>
<%
				else
%>
				<select name=idParentCategory id=idParentCategory size=1>
					<option value=""></option>
<%
					'Get all Categories
					mySQL="SELECT * " _
					    & "FROM   categories " _
					    & "ORDER BY categoryDesc"
					set rs = openRSexecute(mySQL)
					do while not rs.eof
%>
						<option value="<%=rs("idCategory")%>" <%=checkMatch(idParentCategory,rs("idCategory"))%>><%=rs("categoryDesc")%></option>
<%
						rs.movenext
					loop
					call closeRS(rs)
%>
				</select>
<%
				end if
%>
			</td>
			<td>
				<select name=categoryFeatured id=categoryFeatured size=1>
					<option value="N" <%=checkMatch(categoryFeatured,"N")%>>No</option>
					<option value="Y" <%=checkMatch(categoryFeatured,"Y")%>>Yes</option>
				</select>
			</td>
			<td>
				<input type=hidden name=idCategory id=idCategory value="<%=idCategory%>">
				<input type=hidden name=action     id=action     value="edit">
				<input type=submit name=submit1    id=submit1    value="Update">
			</td>
		</tr>
		<tr>
			<td nowrap colspan=4><b>Category HTML</b></td>
		</tr>
		<tr>
			<td nowrap colspan=4><input type=text name=CategoryHTML id=CategoryHTML size=50 maxlength=255 value="<%=server.HTMLEncode(CategoryHTML & "")%>"></td>
		</tr>
	</form>
	</table>
<%
end if

'Add
if action = "add" then
%>
	<span class="textBlockHead">Add Category</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<form method="post" action="SA_cat_exec.asp" name="form1">
		<tr>
			<td nowrap><b>Category Description</b></td>
			<td nowrap><b>Parent Category</b></td>
			<td nowrap><b>Featured?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>
				<input type=text name=categoryDesc id=categoryDesc size=20 maxlength=50>
			</td>
			<td>
				<select name=idParentCategory id=idParentCategory size=1>
					<option value=""></option>
<%
					'Get all Categories
					mySQL="SELECT * " _
					    & "FROM   categories " _
					    & "ORDER BY categoryDesc"
					set rs = openRSexecute(mySQL)
					do while not rs.eof
%>
						<option value="<%=rs("idCategory")%>"><%=rs("categoryDesc")%></option>
<%
						rs.movenext
					loop
					call closeRS(rs)
%>
				</select>
			</td>
			<td>
				<select name=categoryFeatured id=categoryFeatured size=1>
					<option value="N">No</option>
					<option value="Y">Yes</option>
				</select>
			</td>
			<td>
				<input type=hidden name=action  id=action  value="add">
				<input type=submit name=submit1 id=submit1 value="Add">
			</td>
		</tr>
		<tr>
			<td nowrap colspan=4><b>Category HTML</b></td>
		</tr>
		<tr>
			<td nowrap colspan=4><input type=text name=CategoryHTML id=CategoryHTML size=50 maxlength=255></td>
		</tr>
	</form>
	</table>
<%
end if

'Delete
if action = "del" then
%>
	<span class="textBlockHead">Delete Category</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<form method="post" action="SA_cat_exec.asp" name="form2">
		<tr>
			<td nowrap><b>Category Description</b></td>
			<td nowrap><b>Parent Category</b></td>
			<td nowrap><b>Featured?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td><%=CategoryDesc%></td>
			<td><%=ParentCategoryDesc%></td>
			<td><%=CategoryFeatured%></td>
			<td>
				<input type=hidden name=idCategory id=idCategory value="<%=idCategory%>">
				<input type=hidden name=action     id=action     value="del">
				<input type=submit name=submit1    id=submit1    value="Delete">
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
		
		<b>Category Description</b> - Mandatory. A name for the 
		category.<br><br>
		
		<b>Parent Category</b> - Mandatory. The parent category to 
		which a sub-category is linked. A sub-category can also act 
		as the parent category for another sub-category, thereby 
		creating a nested category structure than can be many levels 
		deep.<br><br>
		
		<b>Featured?</b> - Mandatory. This is used to determine if the 
		category will be dynamically displayed in the navigation bars. 
		If you will be creating your own navigation menus, this setting 
		is not really important.<br><br>
		
		<b>Category HTML</b> - Optional. This field can be used to 
		display additional information or images related to a Category. 
		Because you are allowed to use HTML, you can add an image to the 
		Category "tree" display using the IMG tag, or you can display 
		text in bold, etc. There is a 255 character restriction on this 
		field. Some examples are :<br><br>
		
<PRE>
&lt;IMG SRC="../CatImg/image.gif"&gt;
&lt;b&gt;BOLD TEXT&lt;/b&gt;
</PRE>
		
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
	<a href=SA_cat.asp>List Catgeories</a> | 
	<a href=SA_cat_edit.asp?action=edit&recID=<%=idCategory%>>Edit</a> | 
	<a href=SA_cat_edit.asp?action=add>Add</a> | 
	<a href=SA_cat_edit.asp?action=del&recID=<%=idCategory%>>Delete</a> 
	]
<%
end sub
%>

