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
dim mySQL, cn, rs, rs2

'Options
dim idOption
dim optionDescrip
dim priceToAdd
dim weightToAdd
dim taxExempt
dim percToAdd

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
	<b><font size=3>Option Maintenance</font></b>
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
call optTabs("OP")

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
			<a href="SA_opt_edit.asp?action=add">Add New Option</a>
		</td>
	</tr>
</table>

<br>
	
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="listTable">
	<tr>
		<td colspan=8 class="listRowTop">&nbsp;</td>
	</tr>
	<form method="post" action="SA_opt_exec.asp" name="form2" id="form2">
	<tr>
		<td class="listRowHead"><b>ID</b></td>
		<td class="listRowHead"><b>Option Description</b></td>
		<td class="listRowHead" align=right><b>Price</b></td>
		<td class="listRowHead" align=right><b>Perc.</b></td>
		<td class="listRowHead" align=right><b>Weight</b></td>
		<td class="listRowHead" align=center><b>TaxEx</b></td>
		<td class="listRowHead" width="1%"><b>&nbsp;</b></td>
		<td class="listRowHead" width="1%" align=center>
			<input type="checkbox" name="checkAll" value="1" onclick="javascript:CheckAll(this.form);">
		</td>
	</tr>
<%
	rowColor = col2

	'Retrieve all Options
	mySQL="SELECT * " _
	    & "FROM   Options " _
	    & "ORDER BY optionDescrip"
	set rs = openRSexecute(mySQL)
	do while not rs.eof
%>
		<tr>
			<td bgcolor="<%=rowColor%>" valign=top><%=rs("idOption")%></td>
			<td bgcolor="<%=rowColor%>" valign=top><%=rs("optionDescrip")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>" align=right valign=top><%=moneyD(rs("priceToAdd"))%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>" align=right valign=top><%=formatNumber(rs("percToAdd"))%>%&nbsp;</td>
			<td bgcolor="<%=rowColor%>" align=right valign=top><%=rs("weightToAdd")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>" align=center valign=top><%=rs("taxExempt")%>&nbsp;</td>
			<td bgcolor="<%=rowColor%>" valign=top nowrap>
				[ <a href="SA_opt_edit.asp?action=edit&recid=<%=rs("idOption")%>">edit</a> ]
			</td>
			<td align=center valign=top bgcolor="<%=rowColor%>">
				<input type=checkbox name="idOption" id="idOption" value="<%=rs("idOption")%>">
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
		<td colspan=8 align=right nowrap class="listRowBot">
			<input type=hidden name="action" id="action" value="bulkDel">
			Delete Selected Options? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
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
	
	<b>Overview</b> - Options are used to collect additional bits of 
	information from a customer when they order a particular product. 
	There are two ways in which you can do 
	this. First of all, you can display a group of related options in a 
	drop-down select list which will allow the customer to select one 
	of the options in that list. Secondly, you can display a text box 
	into which the customer can enter some text.<br><br>
	
	<b>Example - Drop-Down Select List :</b><br>
	
	<ol>
		<li>Scenario - You want the customer to be able to pick a 
		color from a list of colors when they order a T-Shirt (see 
		example structure below).</li>
		
		<li>Create an option for each color. For example, Red, Green, 
		Blue, etc. Note : On the store's product detail page, the 
		option description will be displayed in the list, along with 
		it's price (if any).</li>
		
		<li>Create a Color option group using the 
		<a href="SA_optGrp.asp">Option Groups</a> function, and add 
		the options you created in the previous step to this 
		option group. Make sure that the option group type 
		is "Drop-Down List". Note : On the store's product detail page, 
		the option group description will be displayed above the 
		drop-down select list as a heading for that list of options.</li>
		
		<li>The final step would be to add the Color option group to 
		each T-Shirt product. This is done using the 
		<a href="SA_prod.asp?resetCookie=1">Product Maintenance</a> 
		function. When you add the option group to a product, you will 
		also have the option of excluding any options in the option 
		group that do not apply to that particular product. This way 
		you don't have to create multiple Color option groups for 
		different products with different color combinations.</li>
		
		<li>The system will now display a drop-down box with a list of 
		Colors on the product detail page. The same process can be 
		used for many other options, such as T-Shirt Size, Gift 
		Wrapping, Technical Support Options, etc.</li>

	</ol>
	
	<b>Example - Text Box :</b><br>
	
	<ol>
		<li>Scenario - You want the customer to enter some text which 
		will be printed on the T-Shirt (see example structure below).</li>
		
		<li>Create a single option and give it a meaningful description. 
		Note : Unlike drop-down list options, this description will 
		not be displayed on the product detail page. Instead, it's merely 
		used to identify that option in the maintenance functions. Like 
		any other option however, you can specify if this option will 
		incur an additional price.</li>
		
		<li>Create a Printed Text option group 
		using the <a href="SA_optGrp.asp">Option Groups</a> function, 
		and add the option you created in the previous step to  
		this option group. Make sure that the option group type is 
		"Text Input". Note : On the store's product detail page, the 
		option group description will be displayed above the text box 
		as a heading for that text box.</li>
		
		<li>The final step would be to add the Printed Text option 
		group to each T-Shirt product. This is done using 
		the <a href="SA_prod.asp?resetCookie=1">Product Maintenance</a> 
		function.</li>
		
		<li>The system will now display a text box on the product detail 
		page. Other good examples would be where you want the customer 
		to enter an engraving for jewellery, a domain name when they 
		order web hosting, etc.</li>

	</ol>
	
	<b>Example Structure :</b><br>
	
	<ul>
		<li>T-Shirt
			<ul>
				<li>Color
					<ul>
						<li><font color=blue>White</font></li>
						<li><font color=blue>Red</font></li>
						<li><font color=blue>Blue</font></li>
					</ul>
				</li>
				<li>Size
					<ul>
						<li><font color=blue>S</font></li>
						<li><font color=blue>M</font></li>
						<li><font color=blue>L</font></li>
						<li><font color=blue>XL</font></li>
					</ul>
				</li>
				<li>Printed Text
					<ul>
						<li><font color=blue>&lt;Text entered by customer&gt;</font></li>
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