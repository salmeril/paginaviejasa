<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Options Maintenance
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

'Options
dim idOption
dim optionDescrip
dim priceToAdd
dim weightToAdd
dim taxExempt
dim percToAdd

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
	<b><font size=3>Option Maintenance</font></b>
	<br><br>
</P>

<%
'Page Tabs
call optTabs("OP")

'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if  action <> "edit" _
and action <> "del"  _
and action <> "add"  then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idOption
if action = "edit" _
or action = "del" then
	idOption = trim(Request.QueryString("recId"))
	if len(idOption) = 0 then
		idOption = trim(Request.Form("recId"))
	end if
	if idOption = "" or not isNumeric(idOption) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option ID.")
	end if
end if

'Get Option Record
if action = "edit" or action = "del" then
	mySQL="SELECT optionDescrip,priceToAdd,weightToAdd," _
		& "       taxExempt,percToAdd " _
	    & "FROM   options " _
	    & "WHERE  idOption = " & idOption
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option ID.")
	else
		optionDescrip = rs("optionDescrip")
		priceToAdd    = rs("priceToAdd")
		weightToAdd   = rs("weightToAdd")
		taxExempt	  = rs("taxExempt")
		percToAdd	  = rs("percToAdd")
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
	<span class="textBlockHead">Edit Option</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td nowrap><b>Option Description</b></td>
			<td nowrap><b>Price</b></td>
			<td nowrap><b>Perc.</b></td>
			<td nowrap><b>Weight</b></td>
			<td nowrap><b>Tax Exempt?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_opt_exec.asp" name="form1">
			<td>
				<input type=text name=optionDescrip id=optionDescrip size=25 maxlength=50 value="<%=optionDescrip%>">
			</td>
			<td>
				<input type=text name=priceToAdd id=priceToAdd size=5 maxlength=10 value="<%=moneyD(priceToAdd)%>">
			</td>
			<td>
				<input type=text name=percToAdd id=percToAdd size=5 maxlength=10 value="<%=formatNumber(percToAdd,2)%>">
			</td>
			<td>
				<input type=text name=weightToAdd id=weightToAdd size=5 maxlength=10 value="<%=weightToAdd%>">
			</td>
			<td>
				<select name=taxExempt id=taxExempt size=1>
					<option value="N" <%=checkMatch(taxExempt,"N")%>>No</option>
					<option value="Y" <%=checkMatch(taxExempt,"Y")%>>Yes</option>
				</select>
			</td>
			<td>
				<input type=hidden name=idOption id=idOption value="<%=idOption%>">
				<input type=hidden name=action   id=action   value="edit">
				<input type=submit name=submit1  id=submit1  value="Update">
			</td>
			</form>
		</tr>
		<tr>
			<td colspan=6>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=6 bgcolor="#dddddd">
				<span class="textBlockHead">Option Groups</span> 
			</td>
		</tr>
		<tr>
			<td colspan=6>
				<table border=0 cellspacing=0 cellpadding=5 bgcolor=#eeeeee>
<%
					'Get Option Groups for this Option
					mySQL="SELECT b.idOptionGroup, b.optionGroupDesc " _
						& "FROM   optionsXref a " _
						& "INNER JOIN optionsGroups b " _
						& "ON     a.idOptionGroup = b.idOptionGroup " _
						& "WHERE  a.idOption = " & idOption
					set rs = openRSexecute(mySQL)
					do while not rs.eof
%>
						<tr>
							<td nowrap><%=rs("optionGroupDesc")%></td>
							<td nowrap>
								<a href="SA_optGrp_edit.asp?action=edit&recID=<%=rs("idOptionGroup")%>">Edit</a>
							</td>
						</tr>
<%
						rs.movenext
					loop
					call closeRS(rs)
%>
					<tr>
						<td><a href="SA_optGrp.asp">Add Option to Option Group</a></td>
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
	<span class="textBlockHead">Add Option</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<form method="post" action="SA_opt_exec.asp" name="form1">
		<tr>
			<td nowrap><b>Option Description</b></td>
			<td nowrap><b>Price</b></td>
			<td nowrap><b>Perc.</b></td>
			<td nowrap><b>Weight</b></td>
			<td nowrap><b>Tax Exempt?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>
				<input type=text name=optionDescrip id=optionDescrip size=30 maxlength=50>
			</td>
			<td>
				<input type=text name=priceToAdd id=priceToAdd size=5 maxlength=10 value="0.00">
			</td>
			<td>
				<input type=text name=percToAdd id=percToAdd size=5 maxlength=10 value="0.00">
			</td>
			<td>
				<input type=text name=weightToAdd id=weightToAdd size=5 maxlength=10 value="0">
			</td>
			<td>
				<select name=taxExempt id=taxExempt size=1>
					<option value="N" selected>No</option>
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
	<span class="textBlockHead">Delete Option</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td nowrap><b>Option Description</b></td>
			<td nowrap><b>Price</b></td>
			<td nowrap><b>Perc.</b></td>
			<td nowrap><b>Weight</b></td>
			<td nowrap><b>Tax Exempt?</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form method="post" action="SA_opt_exec.asp" name="form2">
			<td><%=optionDescrip%></td>
			<td><%=moneyD(priceToAdd)%></td>
			<td><%=formatNumber(percToAdd,2)%></td>
			<td><%=weightToAdd%></td>
			<td><%=taxExempt%></td>
			<td>
				<input type=hidden name=idOption id=idOption value="<%=idOption%>">
				<input type=hidden name=action   id=action   value="del">
				<input type=submit name=submit1  id=submit1  value="Delete">
			</td>
			</form>
		</tr>
	</table>
<%
end if

if action = "edit" or action = "add" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>
	
		<b>Option Description</b> - Mandatory. A name for the 
		option.<br><br>
		
		<b>Price</b> - Optional. You can specify a fixed Price for an 
		option.<br><br>
		
		<b>Percentage</b> - Optional. You can specify that the option 
		Price be calculated as a percentage of the Product Price. If 
		both a fixed price and percentage is entered, the system will 
		use the higher amount. This way you can ensure a "minimum" 
		option price.<br><br>
		
		<b>Weight</b> - Optional. If entered, the weight will be 
		factored into the 
		shipping cost calculation, along with the product's weight. 
		Note that, if a product with 'Free Shipping' set to 'Yes' 
		(see product maintenance) is added to the shopping cart, 
		all related options added to the shopping cart will also 
		be considered free of shipping (ie. weight will be 
		ignored). You must use the same Unit of Weight (Kilogram, 
		Grams, Pounds, etc.) for the option, product and shipping rate 
		weights (see Product & Shipping Maintenance). So, if you 
		enter the weight in "pounds" here, you must also enter the 
		weight in "pounds" for products and shipping rates.<br><br>
		
		<b>Tax Exempt</b> - Mandatory. The Tax Exempt indicator is 
		used when calculating taxes. If set to "Yes", no taxes will 
		be calculated on the price of the option. If set to "No", 
		the option's price will be added into the tax calculation.<br><br>
		
<%
		if action = "edit" then
%>
		<b>Option Groups</b> - This is a list of Option Groups that 
		this Option is currently linked to. To add the Option to 
		another Option Group, use the 
		<a href="SA_optGrp.asp">Option Groups</a> 
		function.
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
	<a href=SA_opt.asp>List Options</a> | 
	<a href=SA_opt_edit.asp?action=edit&recID=<%=idOption%>>Edit</a> | 
	<a href=SA_opt_edit.asp?action=add>Add</a> | 
	<a href=SA_opt_edit.asp?action=del&recID=<%=idOption%>>Delete</a> 
	]
<%
end sub
%>

