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
<%
'Database
dim mySQL, cn, rs

'Categories
dim idCategory
dim categoryDesc
dim idParentCategory
dim categoryFeatured
dim categoryHTML

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
	<b><font size=3>Category Maintenance</font></b>
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

'Check if the Root Category (idParentCategory=0) exists
mySQL="SELECT idCategory " _
    & "FROM   categories " _
    & "WHERE  idParentCategory = 0 "
set rs = openRSexecute(mySQL)
if rs.eof then
	call closeRS(rs)
%>
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<tr><td>
		<br><b>The store requires one "Root" Category.</b><br><br><br>
		Would you like to create a Root Category now? 
		[ <a href="SA_cat_exec.asp?action=root">Yes</a> | <a href="default.asp">No</a> ]<br><br><br>
	</td></tr>
	</table>
	<!--#include file="_INCfooter_.asp"-->
<%
	Response.End
end if
call closeRS(rs)
	
'Retrieve all Categories
mySQL="SELECT a.idCategory, a.categoryDesc, " _
	& "       a.idParentCategory, a.categoryFeatured, " _
	& "      (SELECT b.categoryDesc " _
	& "       FROM   categories b " _
	& "       WHERE  b.idCategory = a.idParentCategory) " _
	& "       AS ParentCategoryDesc " _ 
    & "FROM   categories a " _
    & "ORDER BY a.idParentCategory, a.CategoryDesc"
set rs = openRSexecute(mySQL)
%>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="findTable">
	<tr>
		<td align=left valign=middle nowrap>
			<a href="SA_cat_edit.asp?action=add">Add New Category</a>
		</td>
	</tr>
</table>

<br>
	
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="listTable">
	<tr>
		<td colspan=6 class="listRowTop">&nbsp;</td>
	</tr>
<%
	rowColor = col1
%>
	<form method="post" action="SA_cat_exec.asp" name="form2" id="form2">
	<tr>
		<td class="listRowHead" width="2%"><b>ID</b></td>
		<td class="listRowHead" width="47%"><b>Category</b></td>
		<td class="listRowHead" width="47%" nowrap><b>Parent Category</b></td>
		<td class="listRowHead" width="2%" align=center><b>Feat.</b></td>
		<td class="listRowHead" width="1%">&nbsp;</td>
		<td class="listRowHead" width="1%" align=center>
			<input type="checkbox" name="checkAll" value="1" onclick="javascript:CheckAll(this.form);">
		</td>
	</tr>
<%
	rowColor = col2
	rs.MoveFirst
	do while not rs.eof
%>
		<tr>
			<td bgcolor="<%=rowColor%>"><%=rs("idCategory")%></td>
			<td bgcolor="<%=rowColor%>"><%=rs("categoryDesc")%></td>
<%
			if rs("idParentCategory") = 0 then
%>
				<td bgcolor="<%=rowColor%>">&lt;ROOT&gt;</td>
<%
			elseif isNull(rs("ParentCategoryDesc")) or rs("ParentCategoryDesc") = "" then
%>
				<td bgcolor="<%=rowColor%>"><span class="errMsg">--&gt; None</span></td>
<%
			else
%>
				<td bgcolor="<%=rowColor%>">--&gt; <%=rs("ParentCategoryDesc")%></td>
<%
			end if			
%>
			<td bgcolor="<%=rowColor%>" align=center><%=rs("categoryFeatured")%></td>
			<td bgcolor="<%=rowColor%>" nowrap>
				[ 
				<a href="SA_cat_edit.asp?action=edit&recid=<%=rs("idCategory")%>">edit</a> | 
				<a href="../scripts/prodlist.asp?idCategory=<%=rs("idCategory")%>">test</a> 
				]
			</td>
			<td align=center valign=top bgcolor="<%=rowColor%>">
<%
				if rs("idParentCategory") = 0 then
					Response.Write "&nbsp;"
				else
					Response.Write "<input type=checkbox name=idCategory id=idCategory value=""" & rs("idCategory") & """>"
				end if
%>
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
		<td colspan=6 align=right nowrap class="listRowBot">
			<input type=hidden name="action" id="action" value="bulkDel">
			Delete Selected Categories? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
		</td>
	</tr>
	</form>
<%
	call closeRS(rs)
%>
</table>

<br>

<span class="textBlockHead">Help and Instructions :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr><td>
	
	<b>Example</b> - You can have as many Categories within 
	Categories (Sub-Categories) as you want.
	<br>
	<ul>
		<li><font color=blue>Category 1</font>
			<ul>
				<li>Product 1</li>
				<li>Product 2</li>
			</ul>
		</li>
		<li><font color=blue>Category 2</font>
			<ul>
				<li><font color=blue>Category 3</font>
					<ul>
						<li><font color=blue>Category 4</font>
							<ul>
								<li>Product 3</li>
								<li>Product 4</li>
							</ul>
						</li>
					</ul>
				</li>
			</ul>
		</li>
		<li><font color=blue>Category 5</font></li>
	</ul>

	<b>Step By Step</b> - To create the above category structure, 
	you would create Category 1 first and link it to the Root 
	Category. Then you would create Category 2 and link it to the 
	Root Category. Next you would create Category 3 and link it to 
	Category 2. After that you would create Category 4 and link it 
	to Category 3. Lastly you would create Category 5 and link it to 
	the Root Category. Once the Categories have been set up, you can 
	then start linking Products to their required Categories.<br><br>
	
	<b>Basic Rules</b> - Some of the basic rules regarding Categories 
	are listed below :<br><br>
	
	<ol>
		<li>You can NOT link a Category to another Category that :<br><br>
			<ul>
				<li><b>Has Products linked to it</b>. In the above 
				example, 
				you will not be able to link a Category to Categories 
				1 or 4. You will have to un-link the Products from 
				these Categories first.<br><br>
				
				<li><b>Is already a Sub-Category of the Category</b>. 
				In the 
				example above, you will not be able to link 
				Category 2 to Category 3 as it's already a  
				Sub-Category of Category 2.
			</ul><br>
			
		<li>The Category Description can be anything you want, even 
		if that Description is already used with another Category.<br><br>
		
		<li>Deleting a Category will not automatically delete any 
		Sub-Categories or Products linked to it. These 
		Sub-Categories and Products will merely be un-linked from 
		the deleted Category. Categories not currently linked to a 
		Parent Category will be marked in 
		<span class="errMsg">red</span> on this list and should 
		ideally be deleted also, or re-linked to other Categories. 
		So, in the example above, 
		deleting Category 3 will NOT delete Category 4 (or any of 
		it's Products). However, if you list the Category "tree", 
		the listing will only include Categories 1,2 and 5. Category 
		4 will not be displayed as it's no longer linked to anything.<br><br> 		
		
		<li>There is no physical limitation on the number of 
		Sub-Category levels you can create. It is however adviseable 
		not to create a Category structure that is more than 5 levels 
		"deep" as it may have an impact on your server's performance.
		
	</ol>
		
	<b>Test Category</b> - Click on "test" to see what your Category 
	will look like when it's accessed through your store front. 
	When you Test a Category that has Sub-Categories, you should see 
	a list of all the Sub-Categories. If the Category being tested 
	has Products linked to it, you will see the Product list for that 
	Category (if you don't have any Products linked to the Category 
	yet, you will just see a message to that affect).<br><br>
	
</td></tr>
</table>
<%
'Close the Database Connection
call closedb()
%>

<!--#include file="_INCfooter_.asp"-->