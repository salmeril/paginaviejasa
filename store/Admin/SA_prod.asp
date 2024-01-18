<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Product Maintenance
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

'Products
dim idProduct
dim description
dim descriptionLong
dim details
dim relatedKeys
dim price
dim listPrice
dim imageURL
dim smallImageURL
dim sku
dim stock
dim weight
dim active
dim hotDeal
dim homePage
dim fileName
dim noShipCharge
dim reviewAllow
dim reviewAutoActive

'Work Fields
dim I
dim item
dim count
dim pageSize
dim totalPages
dim showArr
dim sortField

dim curPage
dim showPhrase
dim showField
dim showStart
dim showActive
dim showHome
dim showSpecial
dim showShip
dim showReview

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

'Get showPhrase
showPhrase = Request.Form("showPhrase")					'Form
if len(showPhrase) = 0 then
	showPhrase = Request.QueryString("showPhrase")		'QueryString
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

'Get showActive
showActive = Request.Form("showActive")					'Form
if len(showActive) = 0 then
	showActive = Request.QueryString("showActive")		'QueryString
end if

'Get showHome
showHome = Request.Form("showHome")						'Form
if len(showHome) = 0 then
	showHome = Request.QueryString("showHome")			'QueryString
end if

'Get showSpecial
showSpecial = Request.Form("showSpecial")				'Form
if len(showSpecial) = 0 then
	showSpecial = Request.QueryString("showSpecial")	'QueryString
end if

'Get showShip
showShip = Request.Form("showShip")						'Form
if len(showShip) = 0 then
	showShip = Request.QueryString("showShip")			'QueryString
end if

'Get showReview
showReview = Request.Form("showReview")					'Form
if len(showReview) = 0 then
	showReview = Request.QueryString("showReview")		'QueryString
end if

'Check if a Cookie Reset was requested
if Request.QueryString("resetCookie") = "1" then
	Response.Cookies("ProdSearch").expires = Date() - 30
else
	'Check if a Cookie Recall was requested
	if Request.QueryString("recallCookie") = "1" then
		for each item in Request.Cookies
			if item = "ProdSearch" then
				showArr		= Split(Request.Cookies(item),"*|*")
				curPage		= showArr(0)
				showPhrase	= showArr(1)
				showField	= showArr(2)
				showStart	= showArr(3)
				showActive	= showArr(4)
				showHome	= showArr(5)
				showSpecial	= showArr(6)
				showShip	= showArr(7)
				showReview  = showArr(8)
			end if
		next
	else
		'Save Search Criteria in a Cookie
		Response.Cookies("ProdSearch") = navCookie(curPage)
		Response.Cookies("ProdSearch").expires = Date() + 30
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
	showField = "SKU"
end if

'Check what we will be sorting the results on
sortField = Request.Form("sortField")					'Form
if len(sortField) = 0 then
	sortField = Request.QueryString("sortField")		'QueryString
end if
if len(sortField) = 0 then
	sortField = "SKU"
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Product Maintenance</font></b>
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
			<a href="SA_prod_edit.asp?action=add">Add New Product</a>
		</td>

		<form method="post" action="SA_prod.asp" name="form2">
		<td align=right valign=top nowrap>
			Show Products containing the phrase&nbsp;
			<input type=text name=showPhrase id=showPhrase size=20 maxlength=50 value="<%=showPhrase%>">&nbsp;
			<input type=submit name=submit1 id=submit1 value="Find">
		</td>
		</form>
	</tr>
	
	<tr>
		<form method="post" action="SA_prod.asp" name="form3">
		<td align=right valign=top nowrap colspan=2>
			Show Products where&nbsp;
			<select name=showField id=showField size=1>
				<option value="description" <%=checkMatch(showField,"description")%>>Description</option>
				<option value="SKU"         <%=checkMatch(showField,"SKU")        %>>SKU</option>
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
		<form method="post" action="SA_prod.asp" name="form4">
		<td align=right valign=top nowrap colspan=2>
			<select name=showActive id=showActive size=1>
				<option value=""   <%=checkMatch(showActive,"")  %>>N/A</option>
				<option value="-1" <%=checkMatch(showActive,"-1")%>>Active</option>
				<option value="0"  <%=checkMatch(showActive,"0") %>>InActive</option>
			</select>
			<select name=showHome id=showHome size=1>
				<option value=""   <%=checkMatch(showHome,"")  %>>N/A</option>
				<option value="-1" <%=checkMatch(showHome,"-1")%>>Featured</option>
				<option value="0"  <%=checkMatch(showHome,"0") %>>Not Featured</option>
			</select>
			<select name=showSpecial id=showSpecial size=1>
				<option value=""   <%=checkMatch(showSpecial,"")  %>>N/A</option>
				<option value="-1" <%=checkMatch(showSpecial,"-1")%>>Special</option>
				<option value="0"  <%=checkMatch(showSpecial,"0") %>>Not Special</option>
			</select>
			<select name=showShip id=showShip size=1>
				<option value=""  <%=checkMatch(showShip,"") %>>N/A</option>
				<option value="Y" <%=checkMatch(showShip,"Y")%>>Free Ship.</option>
				<option value="N" <%=checkMatch(showShip,"N")%>>Ship. Charge</option>
			</select>
			<select name=showReview id=showReview size=1>
				<option value=""  <%=checkMatch(showReview,"") %>>N/A</option>
				<option value="Y" <%=checkMatch(showReview,"Y")%>>Review Allow</option>
				<option value="N" <%=checkMatch(showReview,"N")%>>No Reviews</option>
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
	mySQL="SELECT idProduct,sku,description,price," _
		& "       stock,active,homePage,taxExempt, " _
		& "       reviewAllow, " _
		& "      (SELECT (SUM(revRating)/COUNT(revRating)) " _
		& "       FROM    reviews b " _
		& "       WHERE   b.idProduct = a.idProduct " _
		& "       AND     b.revStatus = 'A') " _
		& "       AS      revRating " _
	    & "FROM   Products a " _
	    & "WHERE  1=1 " 'Dummy check to set up conditional checks below
	    
	'Phrase
	if len(showPhrase) > 0 then
		mySQL = mySQL _
			& "AND (description     LIKE '%" & replace(showPhrase,"'","''") & "%'  " _
			& "OR   descriptionLong LIKE '%" & replace(showPhrase,"'","''") & "%'  " _
			& "OR   details         LIKE '%" & replace(showPhrase,"'","''") & "%') " 
	end if
	    
	'Field Start
	if len(showStart) > 0 then
		mySQL = mySQL & "AND " & showField & " LIKE '" & showStart & "%' "
	end if
	    
	'Status
	if len(showActive) > 0 then
		mySQL = mySQL & "AND active = " & showActive & " "
	end if
	
	'Featured
	if len(showHome) > 0 then
		mySQL = mySQL & "AND homepage = " & showHome & " "
	end if
	
	'Specials
	if len(showSpecial) > 0 then
		mySQL = mySQL & "AND hotDeal = " & showSpecial & " "
	end if
	
	'Shipping
	if len(showShip) > 0 then
		mySQL = mySQL & "AND noShipCharge = '" & showShip & "' "
	end if
	
	'Reviews
	if len(showReview) > 0 then
		mySQL = mySQL & "AND reviewAllow = '" & showReview & "' "
	end if
	
	'Sort Order
	mySQL = mySQL & "ORDER BY " & sortField
	    
	set rs = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pageSize)
	if rs.eof then
%>
		<tr>
			<td align=center valign=middle>
				<br>
				<b>No Products matched search criteria.</b>
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
			<td nowrap colspan=3 class="listRowTop">
<%
				call pageNavigation("selectPageTop")
%>
			</td>
			<td colspan=9 align=right class="listRowTop">
				Sort : 
				<select name=sortField id=sortField size=1 onChange="location.href='SA_prod.asp?recallCookie=1&sortField='+this.options[selectedIndex].value">
					<option value="description" <%=checkMatch(sortField,"description")%>>Description</option>
					<option value="SKU"         <%=checkMatch(sortField,"SKU")        %>>SKU</option>
					<option value="price"       <%=checkMatch(sortField,"price")      %>>Price</option>
					<option value="stock"       <%=checkMatch(sortField,"stock")      %>>Stock</option>
				</select>
			</td>
		</tr>
<%
		rowColor = col1
%>
		<form method="post" action="SA_prod_exec.asp" name="form5" id="form5">
		<tr>
			<td class="listRowHead"><b>ID</b></td>
			<td class="listRowHead"><b>SKU</b></td>
			<td class="listRowHead"><b>Description</b></td>
			<td class="listRowHead" align=right><b>Price</b></td>
			<td class="listRowHead" align=right><b>Stock</b></td>
			<td class="listRowHead" align=center><b>FP</b></td>
			<td class="listRowHead" align=center><b>AC</b></td>
			<td class="listRowHead" align=center><b>TE</b></td>
			<td class="listRowHead" align=center><b>RA</b></td>
			<td class="listRowHead" align=center><b>RR</b></td>
			<td class="listRowHead"><b>&nbsp;</b></td>
			<td class="listRowHead" nowrap align=center>
				<input type="checkbox" name="checkAll" value="1" onclick="javascript:CheckAll(this.form);">
			</td>
		</tr>
<%
		rowColor = col2
		do while not rs.eof and count < rs.pageSize
%>
			<tr>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("idProduct")%></td>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("sku")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top><%=rs("description")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top align=right nowrap><%=moneyD(rs("price"))%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top align=right>
<%
					if rs("stock") <= pHideAddStockLevel then
						Response.Write "<font color=red>" & rs("stock") & "</font>"
					else
						Response.Write rs("stock")
					end if
%>
				</td>
				<td bgcolor="<%=rowColor%>" valign=top align=center>
<%
					if rs("homePage") = -1 then
						Response.Write "Y"
					else
						Response.Write "N"
					end if
%>
					&nbsp;
				</td>
				<td bgcolor="<%=rowColor%>" valign=top align=center>
<%
					if rs("active") = -1 then
						Response.Write "Y"
					else
						Response.Write "N"
					end if
%>
					&nbsp;
				</td>
				<td bgcolor="<%=rowColor%>" valign=top align=center><%=rs("taxExempt")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top align=center><%=rs("reviewAllow")%>&nbsp;</td>
				<td bgcolor="<%=rowColor%>" valign=top align=center>
<%
					if isNull(rs("revRating")) then
						Response.Write "-"
					else
						Response.Write round(rs("revRating"),1)
					end if
%>
					&nbsp;
				</td>
				<td bgcolor="<%=rowColor%>" align=right valign=top nowrap>
					[ 
					<a href="SA_prod_edit.asp?action=edit&recid=<%=rs("idProduct")%>">edit</a> 
					]
				</td>
				<td align=middle valign=top bgcolor="<%=rowColor%>">
					<input type=checkbox name="idProduct" id="idProduct" value="<%=rs("idProduct")%>">
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
			<td nowrap colspan=3 class="listRowBot">
<%
				call pageNavigation("selectPageBot")
%>
			</td>
			<td colspan=9 align=right class="listRowBot">
				<input type=hidden name="action" id="action" value="bulkDel">
				Delete Selected Products? <input type=submit name=submit1 id=submit1 value="Yes" onClick="return confirmSubmit()">
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
	
	<b>Overview</b> - Products are the heart of your store. Each 
	product can be linked to one or more Categories, and can have 
	one or more Option Groups linked to them. You can also set 
	Product Discounts for each product, create Product Groups, and 
	much more.<br><br>
	
	<b>Abbreviations</b> - To conserve space in the product listing, 
	some of the headings have been shortened.
	<ul>
		<li>FP - Featured Product? (Yes / No)</li>
		<li>AC - Active? (Yes / No)</li>
		<li>TE - Tax Exempt? (Yes / No)</li>
		<li>RA - Reviews Allowed? (Yes / No)</li>
		<li>RR - Average Review Rating (1 to 5)</li>
	</ul>

	<b>Find Product(s)</b> - You have several options for finding a specific 
	Product.
	<ul>
		<li>You can list all Products which contain a specific phrase 
		in the Short Description, Long Descripton and Product Detail 
		fields.</li>
		
		<li>You can list all Products where the SKU or Description 
		fields start with the specified Alphabetic or Numeric value.</li>
	
		<li>You can list al Products with the Attributes you selected. If 
		you want the search to ignore the setting for an Attribute, select 
		"N/A" for that Attribute. When you select more than one Attribute, 
		the search performs an "AND" operation on the selection. In other 
		words, if you selected "Active" and "Featured", only those Products 
		that are BOTH "Active" AND "Featured" will be listed.</li>
	</ul>
	
	<b>Add Product</b> - Click on "Add New Product" and complete 
	the form as indicated. Once you have added the Product, you will 
	be able to link the product to one or more Categories, add Option 
	Groups, create Product Discounts, set Product Groups, etc.<br><br>

	<b>Edit Product</b> - Change Product information. You can also 
	add or remove the product from a Category, add or remove Option 
	Groups and Product Discounts, and much more.<br><br>
	
	<b>Delete Product</b> - Check the box(es) next to the Product(s) 
	you want to delete and click "Yes" at the bottom.<br><br>
	
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
	            & "&showPhrase="	& server.URLEncode(showPhrase) _
	            & "&showField="		& server.URLEncode(showField) _
	            & "&showStart="		& server.URLEncode(showStart) _
	            & "&showActive="	& server.URLEncode(showActive) _
	            & "&showHome="		& server.URLEncode(showHome) _
	            & "&showSpecial="	& server.URLEncode(showSpecial) _
	            & "&showShip="		& server.URLEncode(showShip) _
	            & "&showReview="	& server.URLEncode(showReview)
end function
'*********************************************************************
'Make Cookie Value for Paging
'*********************************************************************
function navCookie(pageNum)

	navCookie = pageNum		& "*|*" _
	          & showPhrase	& "*|*" _
	          & showField	& "*|*" _
	          & showStart	& "*|*" _
	          & showActive	& "*|*" _
	          & showHome	& "*|*" _
	          & showSpecial	& "*|*" _
	          & showShip	& "*|*" _
	          & showReview
end function
'*********************************************************************
'Display page navigation
'*********************************************************************
sub pageNavigation(formFieldName)
	Response.Write "Page "
	Response.Write "<select onChange=""location.href=this.options[selectedIndex].value"" name=" & trim(formFieldName) & ">"
	for I = 1 to TotalPages
		Response.Write "<option value=""SA_prod.asp" & navQueryStr(I) & "&sortField=" & server.URLEncode(sortField) & """ " & checkMatch(curPage,I) & ">" & I & "</option>" & vbCrlf
	next
	Response.Write "</select>&nbsp;of&nbsp;" & TotalPages & "&nbsp;&nbsp;"
	Response.Write "[&nbsp;"
	if curPage > 1 then
		Response.Write "<a href=""SA_prod.asp" & navQueryStr(curPage-1) & "&sortField=" & server.URLEncode(sortField) & """>Back</a>"
	else
		Response.Write "Back"
	end if
	Response.Write "&nbsp;|&nbsp;"
	if curPage < TotalPages then
		Response.Write "<a href=""SA_prod.asp" & navQueryStr(curPage+1) & "&sortField=" & server.URLEncode(sortField) & """>Next</a>"
	else
		Response.Write "Next"
	end if
	Response.Write "&nbsp;]"
end sub
%>