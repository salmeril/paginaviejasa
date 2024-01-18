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

'OptionsProdEx
dim idOptionsProdEx
dim idOption

'Work Fields
dim action

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Get action
action = trim(lCase(Request.Form("action")))
if len(action) = 0 then
	action = trim(lCase(Request.QueryString("action")))
end if
if  action <> "edit" _
and action <> "del"  _
and action <> "add"  _
and action <> "delgrp" _
and action <> "addgrp" _
and action <> "delcat" _
and action <> "addcat" _
and action <> "delopt" _
and action <> "addopt" _
and action <> "deldisc" _
and action <> "adddisc" _
and action <> "bulkdel" _
and action <> "delpgrp" _
and action <> "addpgrp" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idProduct
if action = "edit" _
or action = "del"  _
or action = "delgrp" _
or action = "addgrp" _
or action = "delcat" _
or action = "addcat" _
or action = "delopt" _
or action = "addopt" _
or action = "deldisc" _
or action = "adddisc" _
or action = "delpgrp" _
or action = "addpgrp" then

	idProduct = trim(Request.Form("idProduct"))
	if len(idProduct) = 0 then
		idProduct = trim(Request.QueryString("idProduct"))
	end if
	if idProduct = "" or not isNumeric(idProduct) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Product ID.")
	else
		idProduct = CLng(idProduct)
	end if

end if

'Get idOptionsProdEx
if action = "delopt" then
	idOptionsProdEx = trim(Request.QueryString("recId"))
	if idOptionsProdEx = "" or not isNumeric(idOptionsProdEx) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid OptionsProdEx ID.")
	else
		idOptionsProdEx = CLng(idOptionsProdEx)
	end if
end if

'Get idOption
if action = "addopt" then
	idOption = trim(Request.QueryString("recId"))
	if idOption = "" or not isNumeric(idOption) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option ID.")
	else
		idOption = CLng(idOption)
	end if
end if

'Get idOptGrpProd
if action = "delgrp" then
	idOptGrpProd = trim(Request.QueryString("recId"))
	if idOptGrpProd = "" or not isNumeric(idOptGrpProd) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid optionsGroupsXref ID.")
	else
		idOptGrpProd = CLng(idOptGrpProd)
	end if
end if

'Get idOptionGroup
if action = "addgrp" then
	'Check Data Format
	idOptionGroup = Request.Form("idOptionGroup")
	if idOptionGroup = "" or not isNumeric(idOptionGroup) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Group.")
	else
		idOptionGroup = CLng(idOptionGroup)
	end if
	'Check Option Group is valid
	mySQL = "SELECT idOptionGroup " _
	      & "FROM   OptionsGroups " _
	      & "WHERE  idOptionGroup = " & idOptionGroup
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Group.")
	end if
	call closeRS(rs)
	'Check Option Group/Product combination does not already exist
	mySQL = "SELECT idOptionGroup " _
	      & "FROM   optionsGroupsXref " _
	      & "WHERE  idOptionGroup = " & idOptionGroup & " " _
	      & "AND    idProduct     = " & idProduct & " "
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Option Group is already linked to this Product.")
	end if
	call closeRS(rs)
end if

'Get idCatProd
if action = "delcat" then
	idCatProd = trim(Request.QueryString("recId"))
	if idCatProd = "" or not isNumeric(idCatProd) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Categories_Products ID.")
	else
		idCatProd = CLng(idCatProd)
	end if
end if

'Get idCategory
if action = "addcat" then
	'Check Data Format
	idCategory = Request.Form("idCategory")
	if idCategory = "" or not isNumeric(idCategory) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Category.")
	else
		idCategory = CLng(idCategory)
	end if
	'Check idCategory is valid
	mySQL = "SELECT idCategory " _
	      & "FROM   Categories " _
	      & "WHERE  idCategory = " & idCategory
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Category.")
	end if
	call closeRS(rs)
	'Check Category/Product combination does not already exist
	mySQL = "SELECT idCategory " _
	      & "FROM   Categories_Products " _
	      & "WHERE  idCategory = " & idCategory & " " _
	      & "AND    idProduct  = " & idProduct & " "
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Product is already linked to this Category.")
	end if
	call closeRS(rs)
	'Check Category not linked to another Category
	mySQL = "SELECT idCategory " _
	      & "FROM   categories " _
	      & "WHERE  idParentCategory = " & idCategory & " "
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Product can anly be linked to a Category which has no other Categories linked to it.")
	end if
	call closeRS(rs)
end if

'Get idDiscProd
if action = "deldisc" then
	'Check Data Format
	idDiscProd = trim(Request.QueryString("recId"))
	if idDiscProd = "" or not isNumeric(idDiscProd) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid DiscProd ID.")
	else
		idDiscProd = CLng(idDiscProd)
	end if
end if

'Get Discount Details
if action = "adddisc" then

	'Get discFromQty
	discFromQty = trim(Request.Form("discFromQty"))
	if len(discFromQty) = 0 or not Isnumeric(discFromQty) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Discount From Quantity.")
	end if
	discFromQty = CDbl(discFromQty)
	
	'Get discToQty
	discToQty = trim(Request.Form("discToQty"))
	if len(discToQty) = 0 or not Isnumeric(discToQty) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Discount To Quantity.")
	end if
	discToQty = CDbl(discToQty)

	'Get discAmt
	discAmt = trim(Request.Form("discAmt"))
	if len(discAmt) = 0 or not Isnumeric(discAmt) then
		discAmt = null
	else
		discAmt = CDbl(discAmt)
	end if
	
	'Get discPerc
	discPerc = trim(Request.Form("discPerc"))
	if len(discPerc) = 0 or not Isnumeric(discPerc) then
		discPerc = null
	else
		discPerc = CDbl(discPerc)
	end if
	
	'Check that either an Amount OR Percentage has been entered
	if (isNull(discAmt) and isNull(discPerc)) or (not (isNull(discAmt)) and not (isNull(discPerc))) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("You must enter a valid Discount Amount OR Percentage.")
	end if
	
	'Check To is Greater or Equeal to From
	if discToQty < discFromQty then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Discount Qty TO can not be less than Qty FROM.")
	end if
	
	'Check discAmt is not negative
	if not isNull(discAmt) then
		if discAmt < 0 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Discount Amount can not have a negative value.")
		end if
	end if
	
	'Check discPerc is not negative or greater than 100
	if not isNull(discPerc) then
		if discPerc < 0 or discPerc > 100 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Discount Percentage can not be less than 0 or greater than 100.")
		end if
	end if
	
	'Check discAmt is not greater that Product Price
	if not isNull(discAmt) then
		if discAmt > CDbl(trim(Request.Form("price"))) then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Discount Amount can not be greater that Product Price.")
		end if
	end if
	
	'Check that Qty FROM does not overlap with an existing discount
	mySQL = "SELECT idDiscProd " _
	      & "FROM   DiscProd " _
	      & "WHERE  idProduct=" & idProduct & " " _
	      & "AND    " & discFromQty & " >= discFromQty " _
	      & "AND    " & discFromQty & " <= discToQty "
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Discount Qty FROM is overlapping with an existing discount record.")
	end if
	call closeRS(rs)
	
	'Check that Qty TO does not overlap with an existing discount
	mySQL = "SELECT idDiscProd " _
	      & "FROM   DiscProd " _
	      & "WHERE  idProduct=" & idProduct & " " _
	      & "AND    " & discToQty & " >= discFromQty " _
	      & "AND    " & discToQty & " <= discToQty "
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Discount Qty TO is overlapping with an existing discount record.")
	end if
	call closeRS(rs)

end if

'Get idProdGroup
if action = "delpgrp" then
	idProdGroup = trim(Request.QueryString("recId"))
	if idProdGroup = "" or not isNumeric(idProdGroup) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid ProdGroup ID.")
	else
		idProdGroup = CLng(idProdGroup)
	end if
end if

'Get Product Group Details
if action = "addpgrp" then

	'Get prodGroupP
	prodGroupP = Request.Form("prodGroupP")
	if prodGroupP = "" or not isNumeric(prodGroupP) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Parent Product Group value.")
	else
		prodGroupP = CLng(prodGroupP)
	end if

	'Get prodGroupC
	prodGroupC = Request.Form("prodGroupC")
	if prodGroupC = "" or not isNumeric(prodGroupC) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Child Product Group value.")
	else
		prodGroupC = CLng(prodGroupC)
	end if
	
	'Check product not already part of a group
	mySQL = "SELECT idProdGroup " _
	      & "FROM   productGroups " _
	      & "WHERE  prodGroupC = " & prodGroupC
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Product is already part of a Product Group.")
	end if
	call closeRS(rs)

end if

'Get Product Details
if action = "edit" or action = "add" then

	'Get sku
	sku = trim(Request.Form("sku"))
	if len(sku) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid SKU.")
	end if
	
	'Get Description
	description	= trim(Request.Form("description"))
	if len(description) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Short Description.")
	end if
	
	'Get Long Description
	descriptionLong	= trim(Request.Form("descriptionLong"))
	if len(descriptionLong) > 250 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Maximum length for Long Description is 250 characters.")
	end if
	
	'Get Details
	details = trim(Request.Form("details"))
	
	'Get Related Keys
	relatedKeys	= trim(Request.Form("relatedKeys"))
	if len(relatedKeys) > 250 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Maximum length for Related Keys are 250 characters.")
	end if
	
	'Get Price
	price = trim(Request.Form("price"))
	if len(price) = 0 or not Isnumeric(price) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Price.")
	end if
	price = Cdbl(price)
	
	'Get listPrice
	listPrice = trim(Request.Form("listPrice"))
	if len(listPrice) = 0 or not Isnumeric(listPrice) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid List Price.")
	end if
	listPrice = Cdbl(listPrice)
	
	'Get stock
	stock = trim(Request.Form("stock"))
	if len(stock) = 0 or not Isnumeric(stock) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Stock Level.")
	end if
	stock = CLng(stock)
	
	'Get weight
	weight = trim(Request.Form("weight"))
	if len(weight) = 0 or not Isnumeric(weight) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Weight.")
	end if
	weight = Cdbl(weight)
	
	'Get active
	active = trim(Request.Form("active"))
	if active <> "-1" and active <> "0" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid 'Active' Indicator.")
	end if
	
	'Get hotDeal
	hotDeal = trim(Request.Form("hotDeal"))
	if hotDeal <> "-1" and hotDeal <> "0" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid 'Special Deal' Indicator.")
	end if
	
	'Get homePage
	homePage = trim(Request.Form("homePage"))
	if homePage <> "-1" and homePage <> "0" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid 'Featured' Indicator.")
	end if
	
	'Get noShipCharge
	noShipCharge = trim(Request.Form("noShipCharge"))
	if noShipCharge <> "Y" and noShipCharge <> "N" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid 'Free Shipping' Indicator.")
	end if
	
	'Get reviewAllow
	reviewAllow = trim(Request.Form("reviewAllow"))
	if reviewAllow <> "Y" and reviewAllow <> "N" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid 'Review Allowed' Indicator.")
	end if
	
	'Get reviewAutoActive
	reviewAutoActive = trim(Request.Form("reviewAutoActive"))
	if reviewAutoActive <> "Y" and reviewAutoActive <> "N" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid 'Review Auto-Active' Indicator.")
	end if
	
	'Get taxExempt
	taxExempt = trim(Request.Form("taxExempt"))
	if taxExempt <> "Y" and taxExempt <> "N" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid 'Tax Exempt' Indicator.")
	end if

	'Get smallImageURL
	smallImageURL = trim(Request.Form("smallImageURL"))
	
	'Get imageURL
	imageURL = trim(Request.Form("imageURL"))
	
	'Get fileName
	fileName = trim(Request.Form("fileName"))
	
	'Check that Price is not less than the largest Discount Amount
	if len(idProduct) > 0 then
		mySQL = "SELECT idDiscProd " _
		      & "FROM   DiscProd " _
		      & "WHERE  idProduct=" & idProduct & " " _
		      & "AND    discAmt > " & Price
		set rs = openRSexecute(mySQL)
		if not rs.eof then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Product Price can not be less that the largest Discount Amount.")
		end if
		call closeRS(rs)
	end if

end if

'ADD
if action = "add" then

	'Add Product
	mySQL = "INSERT INTO Products (" _
		  & "description,descriptionLong,details,relatedKeys,price," _
		  & "listPrice,imageURL,smallImageURL,sku,stock,weight,active," _
		  & "hotDeal,homePage,fileName,noShipCharge,taxExempt," _
		  & "reviewAllow,reviewAutoActive" _
	      & ") VALUES (" _
	      & "'" & replace(description,"'","''")		& "'," _
	      & "'" & replace(descriptionLong,"'","''")	& "'," _
	      & "'" & replace(details,"'","''")			& "'," _
	      & "'" & replace(relatedKeys,"'","''")		& "'," _
	      &       price								& " ," _
	      &       listPrice							& " ," _
	      & "'" & replace(imageURL,"'","''")		& "'," _
	      & "'" & replace(smallImageURL,"'","''")	& "'," _
	      & "'" & replace(sku,"'","''")				& "'," _
	      &       stock								& " ," _
	      &       weight							& " ," _
	      &       active							& " ," _
	      &       hotDeal							& " ," _
	      &       homePage							& " ," _
	      & "'" & replace(fileName,"'","''")		& "'," _
	      & "'" & noShipCharge						& "'," _
	      & "'" & taxExempt							& "'," _
	      & "'" & reviewAllow						& "'," _
	      & "'" & reviewAutoActive					& "' " _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	'Get idProduct of INSERTed Record
	mySQL = "SELECT MAX(idProduct) AS maxIdProduct " _
		  & "FROM   Products "
	set rs = openRSexecute(mySQL)
	idProduct = rs("maxIdProduct")
	call closeRS(rs)
	
	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Product was Added. You may now add Categories, Options, Discounts, etc. below.")
	
end if

'DELETE or BULK DELETE
if action = "del" or action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of idProduct's that will be deleted

	'If just one delete is being performed, we populate just the 
	'first position in the delete array, else we populate the array
	'with a list of all the records that were selected for deletion.
	if action = "del" then
		delArray = split(idProduct)
	else
		delArray = split(Request.Form("idProduct"),",")
	end if

	'Set CursorLocation of the Connection Object to Client
	cn.CursorLocation = adUseClient
	
	'Loop through list of records and delete one by one
	for delI = LBound(delArray) to UBound(delArray)
	
		'BEGIN Transaction
		cn.BeginTrans
	
		'Delete records from Categories_Products
		mySQL = "DELETE FROM Categories_Products " _
		      & "WHERE idProduct = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)
	
		'Delete records from optionsGroupsXref
		mySQL = "DELETE FROM optionsGroupsXref " _
		      & "WHERE idProduct = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)
	
		'Delete records from DiscProd
		mySQL = "DELETE FROM DiscProd " _
		      & "WHERE idProduct = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)
		
		'Delete records from reviews
		mySQL = "DELETE FROM reviews " _
		      & "WHERE idProduct = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)
		
		'Delete records from productGroups
		mySQL = "DELETE FROM productGroups " _
		      & "WHERE prodGroupC = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'Delete any Product Groups with only 1 product
		mySQL = "SELECT COUNT(*) AS prodGroupCount, prodGroupP " _
			  & "FROM   productGroups " _
			  & "GROUP BY prodGroupP"
		set rs = openRSexecute(mySQL)
		do while not rs.EOF
			if rs("prodGroupCount") = 1 then
				mySQL = "DELETE FROM productGroups " _
				      & "WHERE  prodGroupP = " & rs("prodGroupP")
				set rs2 = openRSexecute(mySQL)
			end if
			rs.MoveNext
		loop
		call closeRS(rs)
		
		'Delete records from OptionsProdEx
		mySQL = "DELETE FROM OptionsProdEx " _
		      & "WHERE idProduct = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'Delete records from Products
		mySQL = "DELETE FROM Products " _
		      & "WHERE idProduct = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)
		
		'END Transaction
		cn.CommitTrans
	
	next

	call closedb()
	Response.Redirect "SA_prod.asp?recallCookie=1&msg=" & server.URLEncode("Product(s) were Deleted.")

end if

'EDIT
if action = "edit" then

	'Update Record
	mySQL = "UPDATE Products SET " _
	      & "description='"     & replace(description,"'","''")		& "'," _
	      & "descriptionLong='" & replace(descriptionLong,"'","''")	& "'," _
	      & "details='"		    & replace(details,"'","''")			& "'," _
	      & "relatedKeys='"		& replace(relatedKeys,"'","''")		& "'," _
	      & "price="		    & price								& " ," _
	      & "listPrice="	    & listPrice							& " ," _
	      & "imageURL='"	    & replace(imageURL,"'","''")		& "'," _
	      & "smallImageURL='"   & replace(smallImageURL,"'","''")	& "'," _
	      & "sku='"			    & replace(sku,"'","''")				& "'," _
	      & "stock="		    & stock								& " ," _
	      & "weight="			& weight							& " ," _
	      & "active="			& active							& " ," _
	      & "hotDeal="			& hotDeal							& " ," _
	      & "homePage="			& homePage							& " ," _
	      & "fileName='"		& replace(fileName,"'","''")		& "'," _
	      & "noShipCharge='"	& noShipCharge						& "'," _
	      & "taxExempt='"		& taxExempt							& "'," _
	      & "reviewAllow='"		& reviewAllow						& "'," _
	      & "reviewAutoActive='"& reviewAutoActive					& "' " _
	      & "WHERE  idProduct = " & idProduct
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Product was Updated.")

end if

'ADDCAT
if action = "addcat" then

	'Add Record to Categories_Products
	mySQL = "INSERT INTO Categories_Products (" _
	      & "idProduct,idCategory" _
	      & ") VALUES (" _
	      & idProduct & "," & idCategory _
	      & ")"
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Product was added to Category.")

end if

'DELCAT
if action = "delcat" then

	'Delete Record from Categories_Products
	mySQL = "DELETE FROM Categories_Products " _
	      & "WHERE idCatProd = " & idCatProd
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Product was removed from Category.")

end if

'ADDGRP
if action = "addgrp" then

	'Add Record to optionsGroupsXref
	mySQL = "INSERT INTO optionsGroupsXref (" _
	      & "idProduct,idOptionGroup" _
	      & ") VALUES (" _
	      & idProduct & "," & idOptionGroup _
	      & ")"
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Option Group was added to this Product.")

end if

'DELGRP
if action = "delgrp" then

	'Delete Record in optionsGroupsXref
	mySQL = "DELETE FROM optionsGroupsXref " _
	      & "WHERE idOptGrpProd = " & idOptGrpProd
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Option Group was removed from this Product.")

end if

'DELDISC
if action = "deldisc" then

	'Delete Record in DiscProd
	mySQL = "DELETE FROM DiscProd " _
	      & "WHERE idDiscProd=" & idDiscProd
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Discount was removed from this Product.")

end if

'ADDDISC
if action = "adddisc" then

	'To make the construction of the SQL statement easier, we assign 
	'a text string "null" to numeric values that are null
	if isNull(discAmt) then
		discAmt = "null"
	end if
	if isNull(discPerc) then
		discPerc = "null"
	end if

	'Add Record to DiscProd
	mySQL = "INSERT INTO DiscProd (" _
	      & "discAmt,discFromQty,discToQty,idProduct,discPerc" _
	      & ") VALUES (" _
	      & discAmt			& "," _
	      & discFromQty		& "," _
	      & discToQty		& "," _
	      & idProduct		& "," _
	      & discPerc		& " " _
	      & ")"
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Discount was added to Product.")

end if

'DELPGRP
if action = "delpgrp" then

	'Delete Record in productGroups
	mySQL = "DELETE FROM productGroups " _
	      & "WHERE idProdGroup = " & idProdGroup
	set rs = openRSexecute(mySQL)
	
	'Delete any Product Groups with only 1 product
	mySQL = "SELECT COUNT(*) AS prodGroupCount, prodGroupP " _
		  & "FROM   productGroups " _
		  & "GROUP BY prodGroupP"
	set rs = openRSexecute(mySQL)
	do while not rs.EOF
		if rs("prodGroupCount") = 1 then
			mySQL = "DELETE FROM productGroups " _
			      & "WHERE  prodGroupP = " & rs("prodGroupP")
			set rs2 = openRSexecute(mySQL)
		end if
		rs.MoveNext
	loop
	call closeRS(rs)

	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Product was removed from Product Group.")

end if

'ADDPGRP
if action = "addpgrp" then

	'If necessary, create new Product Group
	if prodGroupP <= 0 then
		set rs = openRSopen("productGroups",adUseServer,adOpenKeySet,adLockOptimistic,adCmdTable,0)
		rs.AddNew
		rs("prodGroupC") = idProduct
		rs.Update
		prodGroupP		 = rs("idProdGroup")
		rs("prodGroupP") = prodGroupP
		rs.Update
		call closeRS(rs)
	end if

	'Add Record to Product Group
	mySQL = "INSERT INTO productGroups (" _
	      & "prodGroupP,prodGroupC" _
	      & ") VALUES (" _
	      & prodGroupP & "," & prodGroupC _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Product was added to Product Group.")

end if

'DELOPT
if action = "delopt" then

	'Delete Record in OptionsProdEx
	mySQL = "DELETE FROM OptionsProdEx " _
	      & "WHERE idOptionsProdEx = " & idOptionsProdEx
	set rs = openRSexecute(mySQL)
	
	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Option was included for this Product.")

end if

'ADDOPT
if action = "addopt" then

	'Add Record to OptionsProdEx
	mySQL = "INSERT INTO OptionsProdEx (" _
	      & "idOption,idProduct" _
	      & ") VALUES (" _
	      & idOption & "," & idProduct _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	call closedb()
	Response.Redirect "SA_prod_edit.asp?action=edit&recID=" & idProduct & "&msg=" & server.URLEncode("Option was excluded for this Product.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_prod.asp?recallCookie=1"

%>
