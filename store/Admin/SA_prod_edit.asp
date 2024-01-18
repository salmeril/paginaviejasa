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
dim mySQL, cn, rs, rs2

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
dim taxExempt
dim reviewAllow
dim reviewAutoActive

'OptionsGroups
dim idOptionGroup
dim optionGroupDesc

'Categories
dim idCategory
dim categoryDesc
dim idParentCategory

'Categories_Products
dim idCatProd

'optionsGroupsXref
dim idOptGrpProd

'DiscProd
dim idDiscProd
dim discAmt
dim discFromQty
dim discToQty
dim discPerc

'productGroups
dim idProdGroup
dim prodGroupP
dim prodGroupC

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
	<b><font size=3>Product Maintenance</font></b>
	<br><br>
</P>

<SCRIPT language="JavaScript">
<!--
/* Limit form input to a certain number of characters */
function formFldTrunc(fldID,fldLen) {
	if (fldID.value.length > fldLen) fldID.value = fldID.value.substr(0,fldLen);
}
/* Show Product Image Popup */
function showPopup(popURL,popName,popWidth,popHeight) {
	var popWin;
	var popAttr = "width="+popWidth+",height="+popHeight+",resizable=1,scrollbars=1";
	popWin = window.open(popURL,popName,popAttr);
	if (popWin.opener == null) popWin.opener = self;
	popWin.focus();
}
//-->
</SCRIPT>

<%
'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if  action <> "edit" _
and action <> "del"  _
and action <> "add"  _
and action <> "copy" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idProduct
if action = "edit" or action = "del" or action = "copy" then
	idProduct = trim(Request.QueryString("recId"))
	if len(idProduct) = 0 then
		idProduct = trim(Request.Form("recId"))
	end if
	if idProduct = "" or not isNumeric(idProduct) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if
	idProduct = CLng(idProduct)
end if

'Get Product Record
if action = "edit" or action = "del" or action = "copy" then
	mySQL="SELECT idProduct,description,descriptionLong," _
		& "       relatedKeys,price,listPrice," _
		& "       imageURL,smallImageURL,sku," _
		& "       stock,weight,active,hotDeal," _
		& "       homePage,fileName,noShipCharge," _
		& "       taxExempt,reviewAllow,reviewAutoActive," _
		& "       details " _
	    & "FROM   Products " _
	    & "WHERE  idProduct = " & idProduct
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	else
		details			= rs("details")
		description		= rs("description")
		descriptionLong	= rs("descriptionLong")
		relatedKeys  	= rs("relatedKeys")
		price			= rs("price")
		listPrice		= rs("listPrice")
		imageURL		= rs("imageURL")
		smallImageURL	= rs("smallImageURL")
		sku				= rs("sku")
		stock			= rs("stock")
		weight			= rs("weight")
		active			= rs("active")
		hotDeal			= rs("hotDeal")
		homePage		= rs("homePage")
		fileName		= rs("fileName")
		noShipCharge	= rs("noShipCharge")
		taxExempt		= rs("taxExempt")
		reviewAllow		= rs("reviewAllow")
		reviewAutoActive= rs("reviewAutoActive")
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
	<span class="textBlockHead">Edit Product</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td colspan=4 bgcolor="#dddddd">
				<span class="textBlockHead">Product Detail</span>
			</td>
		</tr>
<%
		'Show general product info form
		call prodGeneralInfo()
%>
		<tr><td colspan=4>&nbsp;</td></tr>
		<tr>
			<td colspan=4 bgcolor="#dddddd">
				<span class="textBlockHead">Product Group</span>
			</td>
		</tr>
		<tr>
			<td colspan=4 align=left>
				<table border=0 cellspacing=0 cellpadding=5 bgcolor=#eeeeee>
<%
					'Check if product is part of a Product Group
					mySQL="SELECT prodGroupP " _
						& "FROM   productGroups " _
						& "WHERE  prodGroupC  = " & idProduct
					set rs = openRSexecute(mySQL)
					if not rs.EOF then
						prodGroupP = rs("prodGroupP")
%>
						<tr>
							<td nowrap colspan=2><i>Click 'Remove' to remove this product from the Product Group :</i></td>
						</tr>
<%
					else
						prodGroupP = -1
%>
						<tr>
							<td nowrap colspan=2><i>Select a product to group with this product :</i></td>
						</tr>
<%
					end if
					call closeRS(rs)

					'Display products for this Product Group
					mySQL="SELECT a.idProdGroup, a.prodGroupC," _
					    & "       b.description " _
						& "FROM   productGroups a, products b " _
						& "WHERE  a.prodGroupP = " & prodGroupP & " " _
						& "AND    b.idProduct  = a.prodGroupC " _
						& "ORDER BY b.description "
					set rs = openRSexecute(mySQL)
					do while not rs.eof
%>
						<tr>
							<td nowrap><%=rs("description")%></td>
							<td nowrap>
<%
								if rs("prodGroupC") = idProduct then
%>
								<a href="SA_prod_exec.asp?action=delPgrp&idProduct=<%=idProduct%>&recId=<%=rs("idProdGroup")%>">Remove</a>
<%
								else
%>
								<a href="SA_prod_edit.asp?action=edit&recId=<%=rs("prodGroupC")%>">Edit</a>
<%
								end if
%>
							</td>
						</tr>
<%
						rs.movenext
					loop
					call closeRS(rs)
%>
					<tr>
						<form method="post" action="SA_prod_exec.asp" name="prodGrpAdd">
						<td>
							<select name=prodGroupC id=prodGroupC size=1>
								<option value="">-- Select --</option>
<%
								'Display list of products 
								mySQL="SELECT a.idProduct, a.description " _
								    & "FROM   products a " _
								    & "WHERE  a.idProduct <> " & idProduct & " " _
								    & "AND    NOT EXISTS " _
								    & "      (SELECT b.idProdGroup " _
								    & "       FROM   productGroups b " _
								    & "       WHERE  b.prodGroupC = a.idProduct) " _
								    & "ORDER BY a.description "

								set rs = openRSexecute(mySQL)
								do while not rs.eof
									Response.Write "<option value=""" & rs("idProduct") & """>" & rs("description") & "</option>"
									rs.movenext
								loop
								call closeRS(rs)
%>
							</select>
						</td>
						<td>
							<input type=hidden name=idProduct  id=idProduct  value="<%=idProduct%>">
							<input type=hidden name=prodGroupP id=prodGroupP value="<%=prodGroupP%>">
							<input type=hidden name=action     id=action     value="addPgrp">
							<input type=submit name=submit1    id=submit1    value=" Add ">
						</td>
						</form>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan=4 bgcolor="#dddddd">
				<span class="textBlockHead">Categories</span>
			</td>
		</tr>
		<tr>
			<td colspan=4 align=left>
				<table border=0 cellspacing=0 cellpadding=5 bgcolor=#eeeeee>
<%
					'Get Category Record(s) for this Product
					mySQL="SELECT a.idCatProd, b.categoryDesc " _
						& "FROM   Categories_Products a " _
						& "INNER JOIN Categories b " _
						& "ON     a.idCategory = b.idCategory " _
						& "WHERE  a.idProduct = " & idProduct
					set rs = openRSexecute(mySQL)
					do while not rs.eof
%>
						<tr>
							<td nowrap><%=rs("categoryDesc")%></td>
							<td nowrap><a href="SA_prod_exec.asp?action=delCat&idProduct=<%=idProduct%>&recId=<%=rs("idCatProd")%>">Remove</a></td>
						</tr>
<%
						rs.movenext
					loop
					call closeRS(rs)
%>
					<tr>
						<form method="post" action="SA_prod_exec.asp" name="form3">
						<td>
							<select name=idCategory id=idCategory size=1>
								<option value="">-- Select --</option>
<%
								mySQL="SELECT a.idCategory, a.categoryDesc " _
								    & "FROM   Categories a " _
								    & "WHERE NOT EXISTS " _
								    & "      (SELECT b.idCategory " _
								    & "       FROM   categories b " _
								    & "       WHERE  a.idCategory = b.idParentCategory) " _
								    & "ORDER BY a.categoryDesc "
								set rs = openRSexecute(mySQL)
								do while not rs.eof
									Response.Write "<option value=""" & rs("idCategory") & """>" & rs("categoryDesc") & "</option>"
									rs.movenext
								loop
								call closeRS(rs)
%>
							</select>
						</td>
						<td>
							<input type=hidden name=idProduct id=idProduct value="<%=idProduct%>">
							<input type=hidden name=action    id=action    value="addCat">
							<input type=submit name=submit1   id=submit1   value=" Add ">
						</td>
						</form>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan=4 bgcolor="#dddddd">
				<span class="textBlockHead">Option Groups</span>
			</td>
		</tr>
		<tr>
			<td colspan=4 align=left>
				<table border=0 cellspacing=0 cellpadding=5 bgcolor=#eeeeee>
<%
					'Get Option Groups for this Product
					mySQL="SELECT a.idOptGrpProd, b.idOptionGroup, " _
					    & "       b.optionGroupDesc " _
						& "FROM   optionsGroupsXref a " _
						& "INNER JOIN optionsGroups b " _
						& "ON     a.idOptionGroup = b.idOptionGroup " _
						& "WHERE  a.idProduct = " & idProduct & " " _
						& "ORDER BY b.optionGroupDesc "
					set rs = openRSexecute(mySQL)
					do while not rs.eof
%>
						<tr>
							<td nowrap><%=rs("optionGroupDesc")%></td>
							<td nowrap><a href="SA_prod_exec.asp?action=delGrp&idProduct=<%=idProduct%>&recId=<%=rs("idOptGrpProd")%>">Remove Group</a></td>
						</tr>
<%
						'Get Options for Option Group
						mySQL = "SELECT a.idOption, a.optionDescrip, " _
							  & "      (SELECT c.idOptionsProdEx " _
							  & "       FROM   OptionsProdEx c " _
							  & "       WHERE  c.idOption = a.idOption " _
							  & "       AND    c.idProduct = " & idProduct & ") " _
							  & "       AS     idOptionsProdEx " _
						      & "FROM   options a, optionsXref b " _
						      & "WHERE  a.idOption = b.idOption " _
						      & "AND    b.idOptionGroup = " & rs("idOptionGroup")
						set rs2 = openRSexecute(mySQL)
						do while not rs2.eof
							if isNull(rs2("idOptionsProdEx")) then
%>
							<tr>
								<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;<img src="x_Y.gif">&nbsp;<i><%=rs2("optionDescrip")%></i></td>
								<td nowrap><a href="SA_prod_exec.asp?action=addOpt&idProduct=<%=idProduct%>&recId=<%=rs2("idoption")%>">Exclude</a></td>
							</tr>
<%
							else
%>
							<tr>
								<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;<img src="x_N.gif">&nbsp;<i><%=rs2("optionDescrip")%></i></td>
								<td nowrap><a href="SA_prod_exec.asp?action=delOpt&idProduct=<%=idProduct%>&recId=<%=rs2("idOptionsProdEx")%>">Include</a></td>
							</tr>
<%
							end if
							rs2.movenext
						loop
						call closeRS(rs2)

						rs.movenext
					loop
					call closeRS(rs)
%>
					<tr>
						<form method="post" action="SA_prod_exec.asp" name="form2">
						<td>
							<select name=idOptionGroup id=idOptionGroup size=1>
								<option value="">-- Select --</option>
<%
								mySQL="SELECT idOptionGroup, " _
								    & "       optionGroupDesc " _
								    & "FROM   OptionsGroups " _
								    & "ORDER BY optionGroupDesc"
								set rs = openRSexecute(mySQL)
								do while not rs.eof
									Response.Write "<option value=""" & rs("idOptionGroup") & """>" & rs("optionGroupDesc") & "</option>"
									rs.movenext
								loop
								call closeRS(rs)
%>
							</select>
						</td>
						<td>
							<input type=hidden name=idProduct id=idProduct value="<%=idProduct%>">
							<input type=hidden name=action    id=action    value="addGrp">
							<input type=submit name=submit1   id=submit1   value="Add Group">
						</td>
						</form>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan=4 bgcolor="#dddddd">
				<span class="textBlockHead">Discounts</span>
			</td>
		</tr>
		<tr>
			<td colspan=4 align=left>
				<table border=0 cellspacing=0 cellpadding=5 bgcolor=#eeeeee>
					<tr>
						<td><b>Qty From</b></td>
						<td><b>Qty To</b></td>
						<td><b>Amount (ea.)</b></td>
						<td><b>Perc. (ea.)</b></td>
						<td><b>&nbsp;</b></td>
					</tr>
<%
					'Get Discounts for this Product
					mySQL="SELECT idDiscProd,discAmt," _
					    & "       discFromQty,discToQty," _
					    & "       discPerc " _
						& "FROM   DiscProd " _
						& "WHERE  idProduct=" & idProduct & " " _
						& "ORDER BY discFromQty "
					set rs = openRSexecute(mySQL)
					do while not rs.eof
%>
						<tr>
							<td nowrap><%=rs("discFromQty")%></td>
							<td nowrap><%=rs("discToQty")%></td>
							<td nowrap>
<%
								if isNull(rs("discAmt")) then
									Response.Write "-"
								else
									Response.Write moneyD(rs("discAmt"))
								end if
%>
							</td>
							<td nowrap>
<%
								if isNull(rs("discPerc")) then
									Response.Write "-"
								else
									Response.Write rs("discPerc") & "%"
								end if
%>
							</td>
							<td nowrap><a href="SA_prod_exec.asp?action=delDisc&idProduct=<%=idProduct%>&recId=<%=rs("idDiscProd")%>">Remove</a></td>
						</tr>
<%
						rs.movenext
					loop
					call closeRS(rs)
%>
					<tr>
						<form method="post" action="SA_prod_exec.asp" name="discAddForm">
						<td nowrap><input type=text name=discFromQty id=discFromQty size=9 maxlength=9></td>
						<td nowrap><input type=text name=discToQty   id=discToQty   size=9 maxlength=9></td>
						<td nowrap><input type=text name=discAmt     id=discAmt     size=9 maxlength=9></td>
						<td nowrap><input type=text name=discPerc    id=discPerc    size=9 maxlength=9></td>
						<td nowrap>
							<input type=hidden name=idProduct id=idProduct value="<%=idProduct%>">
							<input type=hidden name=price     id=price     value="<%=price%>">
							<input type=hidden name=action    id=action    value="addDisc">
							<input type=submit name=submit1   id=submit1   value=" Add ">
						</td>
						</form>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<%
end if

'Add and Copy
if action = "add" or action = "copy" then
%>
	<span class="textBlockHead">Add Product</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td colspan=4 bgcolor="#dddddd">
				<span class="textBlockHead">Product Detail</span>
			</td>
		</tr>
<%
		'Show general product info form
		call prodGeneralInfo()
%>
	</table>
<%
end if

'Delete
if action = "del" then
%>
	<span class="textBlockHead">Delete Product</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<tr>
			<td align=right nowrap><b>SKU</b></td>
			<td align=left><%=SKU%></td>
		</tr>
		<tr>
			<td align=right nowrap><b>Short Description</b></td>
			<td align=left><%=description%></td>
		</tr>
		<tr>
			<td align=right nowrap><b>Long Description</b></td>
			<td align=left><%=descriptionLong%></td>
		</tr>
		<tr>
			<td align=right nowrap><b>Product Details</b></td>
			<td align=left><%=details%></td>
		</tr>
		<tr>
			<form method="post" action="SA_prod_exec.asp" name="form4">
			<td colspan=2 align=left>
				<br>
				<input type=hidden name=idProduct id=idProduct value="<%=idProduct%>">
				<input type=hidden name=action    id=action    value="del">
				<input type=submit name=submit1   id=submit1   value="Delete Product">
			</td>
			</form>
		</tr>
	</table>
<%
end if

if action = "edit" or action = "add" or action = "copy" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>

	<b>SKU</b> - Mandatory. The SKU number/code of the Product. This may 
	be any combination of AlphaNumeric characters.<br><br>
	
	<b>Short Description</b> - Mandatory. The Short Description is 
	displayed on all Product list and detail pages. It is also saved 
	with the Order. You may enter HTML tags into the Short Description 
	but it's advised to stick to simple text formatting tags such as 
	font color, size, underline, etc.<br><br>
	
	<b>Long Description</b> - Optional. The Long Description is 
	displayed on Product list pages only. You may enter HTML tags into 
	the Long Description but it's advised to stick to simple text 
	formatting tags such as font color, size, underline, etc.<br><br>
	
	<b>Product Details</b> - Optional. You may enter as much text as you 
	want into the Product Details field, including HTML tags, etc. The 
	Product Details field is displayed on the Product Detail page only. 
	This field is typically used to give en expanded explanation of the 
	product itself. You may even want to enter some information regarding 
	warranties, or links to reviews, etc.<br><br>
	
	<b>Related Keys</b> - Optional. This is a way to link different 
	products together. Any keyword(s) entered into this field will be 
	used to create a "See Related Products" link on the Product Details 
	page. When the Customer clicks on that link, a search of the 
	Product file will be performed, returning any Products that match 
	the keyword(s) entered into this field. So, if you want to link 
	together books by the same author, enter the author's name into 
	this field. If you want to link together products with no tangible 
	relation, enter their SKU numbers into this field. Note : Keywords 
	must be seperated by a space, maximum of 250 characters allowed.<br><br>
	
	<b>Price</b> - Mandatory. The Price of the item. This is the amount 
	which will be added to the Shopping Cart, and is used in all Order 
	calculations, including shipping. Although the Price is 
	mandatory, you may enter a Price of "0.00" which would effectively 
	mean that the Product is free.<br><br>
	
	<b>List Price</b> - Mandatory. The List Price of the item. The amount 
	entered into this field will be used to calculate markdowns and 
	display a "You Save..." message on the Product pages. The List Price 
	is usually equal to or more than the Price itself. It's not used in 
	any Order calculations, but is used on the Order List/View pages for 
	display purposes.<br><br>
	
	<b>Items in Stock</b> - Mandatory. The stock level is automatically 
	reduced when the Order is marked as "Paid", "Complete" or "Shipped". 
	The value in this field can also used to check if the item can still 
	be ordered when the stock level drops below a certain number. This 
	behaviour is determined by the "Out of stock level" store configuration 
	setting.<br><br>
	
	<b>Weight</b> - Mandatory. This value is not displayed to the Customer. 
	It's only purpose is to calculate Shipping Charges based on Product 
	Weight. If Shipping for this item must always be based on Price, set 
	the Weight to "0". Downloadable items should 
	also carry a Weight of "0". Important : You MUST use the same Unit of 
	Weight (Kilogram, Grams, Pounds, etc.) for both the Product Weight and 
	the Shipping Rate Weight (see Shipping Maintenance). So, if you enter 
	the Weight in "Pounds" here, you must also enter the Weight in "Pounds" 
	on the Shipping Rates table.<br><br>
	
	<b>Active?</b> - Mandatory. If set to Yes, the product will be 
	displayed in all the relevant Product lists, and the Customer will be 
	able to Order the Product. If set to No, the Product will not be 
	visible to the Customer.<br><br>
	
	<b>Special Deal?</b> - Mandatory. Indicates whether you have a Special 
	running on this Product. It doesn't have any significant use other than 
	grouping together Products which are currently on Sale. The value in 
	this field does not automatically change the Price. You will still 
	have to reduce the Price to reflect your Sale Price.<br><br>
	
	<b>Featured?</b> - Mandatory. Indicates whether this Item must be 
	displayed on your Home Page or not. If there are no featured Items 
	in the Database, the system will display whatever is in 
	"_INCright_.asp".<br><br>
	
	<b>Free Shipping?</b> - Mandatory. If set to Yes, this item will not 
	be taken into account when calculating Shipping Rates, regardless of 
	Price or Weight. If set to No, this item will be factored into the 
	Shipping Rate calculation. Downloadable items would 
	typically have this field set to Yes. Or you may choose to provide 
	Free Shipping on a Product as an incentive.<br><br>
	
	<b>Allow Reviews?</b> - Mandatory. If set to Yes, your customers 
	will be able to rate this product and write a review for it.<br><br>
	
	<b>Review Auto-Active?</b> - Mandatory. If set to Yes, the review 
	will automatically be activated when it's entered by a customer and 
	will therefore be immediately viewable. If set to No, the review 
	will be stored in the database and will only be viewable when you 
	activate it.<br><br>
	
	<b>Tax Exempt?</b> - Mandatory. If set to Yes, this product will NOT 
	be taken into account when calculating taxes for the order.<br><br>
	
	<b>Small Image</b> - Optional. If you have a thumbnail image of the 
	Product in your Product Images directory on your web server you can 
	specify the filename of the image here. The system will use this 
	filename, along with the Product Images directory path you specified 
	in your configuration (currently <font color=red><%=pImagesDir%></font>) 
	to display the image to the user on the Product List pages. If you 
	specify a Small Image, but not a Large Image below, the Small Image 
	will also be displayed on the Product Detail View page.<br><br>
	
	<b>Large Image</b> - Optional. The same rules apply as for "Small Image" 
	above, except that (if entered) this image will be displayed on 
	the Product Detail View page. If Large Image is left empty, the system 
	will display the Small Image in it's place. If you also omitted the 
	Small Image, the Customer will see a "No Image" message. Both Small and 
	Large Images are stored in the same directory on the web server.<br><br>
	
	<b>File to Download</b> - Optional. If this Product is a downloadable 
	Product, you must specify the name of the file that the Customer will 
	download here. The system will look in the directory you specified 
	in your configuration for this file. When the Order is confirmed 
	paid, the Customer will be notified, and they will be able to log on 
	to their Account and download the file. If the Price of the item is 
	"0.00", the Customer will be able to download the file immediately 
	after saving the Order.<br><br>
<%
	if action = "edit" then
%>
	<b>Product Groups</b> - Products that are similar can be linked to 
	each other via Product Groups. This way, if the customer views a 
	product's detail, and that product belongs to a particular Product 
	Group, the system will automatically display the other products 
	belonging to the same Product Group on the Product Detail page. A 
	product can belong to only one Product Group at a time.<br><br>
	
	<b>Categories</b> - A Product can be linked to one or more Categories. 
	If no Categories are specified and the Product Status is Active, the 
	Product is still accessible via the Search. If you link this Product 
	to a Category, it will appear as part of that Category on the Product 
	List pages in addition to being accessible through the Search. Categories 
	need to be set up in <a href="SA_cat.asp">Category Maintenance</a> first 
	before you can link a Product to them.<br><br>
	
	<b>Option Groups</b> - A Product can be linked to one or more Option 
	Groups. If the user selects an Option which has a Price, the Option 
	price will be added to the Product Price. You also have the option to 
	exclude individual Options for a specific Product. This way you 
	can create an Option Group called "Colors" with all the colors used 
	in your store, and then pick only the colors from the Option Group 
	that apply to a particular Product. Option Groups need to be set 
	up in <a href="SA_optGrp.asp">Option Group Maintenance</a> before you can 
	link them to a Product. Note that removing an Option Group from the 
	Product does not delete the Option Group entirely, it merely removes 
	it for the specific Product.<br><br>
	
	<b>Discounts</b> - You can specify discounts that will automatically 
	be applied to the item's price when the quantity on the order 
	matches the range specified in the "Qty From" and "Qty To" fields. 
	You can specify an Amount or Percentage per item as 
	a discount. If you specify a fixed amount it can not be greater 
	than the product's price.
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
'Display General product info form
'*********************************************************************
sub prodGeneralInfo()
%>
	<form method="post" action="SA_prod_exec.asp" name="prodForm">
	<tr>
		<td align=right nowrap><b>SKU</b></td>
<%
		if action = "edit" then
%>
			<td align=left colspan=3><input type=text name=SKU id=SKU size=16 maxlength=16 value="<%=SKU%>"></td>
<%
		else
%>
			<td align=left colspan=3><input type=text name=SKU id=SKU size=16 maxlength=16></td>
<%
		end if
%>
	</tr>
	<tr>
		<td align=right nowrap><b>Short Description</b></td>
		<td align=left colspan=3><input type=text name=description id=description size=50 maxlength=50 value="<%=server.HTMLEncode(description & "")%>"></td>
	</tr>
	<tr>
		<td align=right nowrap><b>Long Description</b><br>(Max 250 Chars.)</td>
		<td align=left colspan=3><textarea name=descriptionLong cols=45 rows=4 onKeyPress="formFldTrunc(this,249)"><%=server.HTMLEncode(descriptionLong & "")%></textarea></td>
	</tr>
	<tr>
		<td align=right nowrap><b>Product Details</b><br>(Unlimited Chars.)</td>
		<td align=left colspan=3><textarea name=details cols=45 rows=6><%=server.HTMLEncode(details & "")%></textarea></td>
	</tr>
	<tr>
		<td align=right nowrap><b>Related Keys</b></td>
		<td align=left colspan=3><input type=text name=relatedKeys id=relatedKeys size=50 maxlength=250 value="<%=server.HTMLEncode(relatedKeys & "")%>"></td>
	</tr>
	<tr>
		<td align=right nowrap><b>Price</b></td>
		<td align=left><input type=text name=price id=price size=10 maxlength=10 value="<%=moneyD(price)%>"></td>
		<td align=right nowrap><b>List Price</b></td>
		<td align=left><input type=text name=listPrice id=listPrice size=10 maxlength=10 value="<%=moneyD(listPrice)%>"></td>
	</tr>
	<tr>
		<td align=right nowrap><b>Items in Stock</b></td>
		<td align=left><input type=text name=stock id=stock size=9 maxlength=9 value="<%=stock%>"></td>
		<td align=right nowrap><b>Weight</b></td>
		<td align=left><input type=text name=weight id=weight size=10 maxlength=10 value="<%=weight%>"></td>
	</tr>
	<tr>
		<td align=right nowrap><b>Active?</b></td>
		<td align=left>
			<select name=active id=active size=1>
				<option value="0"  <%=checkMatch(active,"0") %>>No</option>
				<option value="-1" <%=checkMatch(active,"-1")%>>Yes</option>
			</select>
		</td>
		<td align=right nowrap><b>Special Deal?</b></td>
		<td align=left>
			<select name=hotDeal id=hotDeal size=1>
				<option value="0"  <%=checkMatch(hotDeal,"0") %>>No</option>
				<option value="-1" <%=checkMatch(hotDeal,"-1")%>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align=right nowrap><b>Featured?</b></td>
		<td align=left>
			<select name=homePage id=homePage size=1>
				<option value="0"  <%=checkMatch(homePage,"0") %>>No</option>
				<option value="-1" <%=checkMatch(homePage,"-1")%>>Yes</option>
			</select>
		</td>
		<td align=right nowrap><b>Free Shipping?</b></td>
		<td align=left>
			<select name=noShipCharge id=noShipCharge size=1>
				<option value="N" <%=checkMatch(noShipCharge,"N")%>>No</option>
				<option value="Y" <%=checkMatch(noShipCharge,"Y")%>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align=right nowrap><b>Allow Reviews?</b></td>
		<td align=left>
			<select name=reviewAllow id=reviewAllow size=1>
				<option value="N" <%=checkMatch(reviewAllow,"N")%>>No</option>
				<option value="Y" <%=checkMatch(reviewAllow,"Y")%>>Yes</option>
			</select>
		</td>
		<td align=right nowrap><b>Review Auto-Active?</b></td>
		<td align=left>
			<select name=reviewAutoActive id=reviewAutoActive size=1>
				<option value="N" <%=checkMatch(reviewAutoActive,"N")%>>No</option>
				<option value="Y" <%=checkMatch(reviewAutoActive,"Y")%>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align=right nowrap><b>Tax Exempt?</b></td>
		<td align=left colspan=3>
			<select name=taxExempt id=taxExempt size=1>
				<option value="N" <%=checkMatch(taxExempt,"N")%>>No</option>
				<option value="Y" <%=checkMatch(taxExempt,"Y")%>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align=right nowrap><b>Small Image</b></td>
		<td align=left colspan=3 nowrap>
			<input type=text name=smallImageURL id=smallImageURL size=18 maxlength=50 value="<%=smallImageURL%>"> 
			<input type=button value="Browse" onClick="showPopup('SA_prodImg.asp?upd=SI','popImgS',500,450)">
		</td>
	</tr>
	<tr>
		<td align=right nowrap><b>Large Image</b></td>
		<td align=left colspan=3 nowrap>
			<input type=text name=imageURL id=imageURL size=18 maxlength=50 value="<%=imageURL%>"> 
			<input type=button value="Browse" onClick="showPopup('SA_prodImg.asp?upd=LI','popImgL',500,450)">
		</td>
	</tr>
	<tr>
		<td align=right nowrap><b>File to Download</b></td>
		<td align=left colspan=3>
			<input type=text name=fileName id=fileName size=18 maxlength=250 value="<%=fileName%>">
			<input type=button value="Browse" onClick="showPopup('SA_prodImg.asp?upd=DL','popDwnl',500,450)">
		</td>
	</tr>
	<tr>
		<td colspan=4 align=center>
			<br>
<%
			if action = "edit" then
%>
			<input type=hidden name=idProduct id=idProduct value="<%=idProduct%>">
			<input type=hidden name=action    id=action    value="edit">
			<input type=submit name=submit1   id=submit1   value="Update Product">
<%
			else
%>
			<input type=hidden name=action    id=action    value="add">
			<input type=submit name=submit1   id=submit1   value="Add Product">
<%
			end if
%>
		</td>
	</tr>
	</form>
<%
end sub
'*********************************************************************
'Create Navigation Links
'*********************************************************************
sub maintNavLinks()
%>
	[ 
	<a href=SA_prod.asp?recallCookie=1>List Products</a> | 
	<a href=SA_prod_edit.asp?action=edit&recid=<%=idProduct%>>Edit</a> | 
	<a href=SA_prod_edit.asp?action=del&recid=<%=idProduct%>>Delete</a> | 
	<a href=SA_prod_edit.asp?action=copy&recid=<%=idProduct%>>Copy</a> | 
	<a href="../scripts/prodview.asp?idProduct=<%=idProduct%>">Test</a> 
	]
<%
end sub
%>