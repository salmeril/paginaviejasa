<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : This script handles all the shopping cart functions namely...
'          : - Add item to cart
'          : - Delete item from shopping cart
'          : - Recalculate shopping cart totals
'          : - View shopping cart
' Product  : CandyPress eCommerce Storefront
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
%>
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="_INCconfig_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="_INCappFunctions_.asp"-->  
<%
'Work Fields
dim action		 'What type of action the script must take
dim errorMsg	 'Error message when adding, deleting, updating
dim errorMsgDisc 'Error message when updating order discount
dim f, i		 'Indexes
dim newQuantity	 'New item quantity (used when updating qty)
dim oldQuantity	 'Old Item quantity (used when updating qty)

'cartHead
dim discCode
dim discPerc
dim discTotal

'cartRows
dim IDCartRow
dim IDProduct
dim SKU
dim quantity
dim unitPrice
dim unitWeight
dim description
dim taxExempt
dim discAmt

'cartRowsOptions
dim idOption
dim optionPrice
dim optionWeight
dim optionDescrip
dim optionTaxExempt

'products
dim stock

'options
dim priceToAdd
dim percToAdd

'DiscOrder
dim idDiscOrder
dim discFromAmt
dim discToAmt

'DiscProd
dim idDiscProd
dim discFromQty
dim discToQty

'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2

'Session
dim idOrder
dim idCust

'*************************************************************************

'Open Database Connection
call openDb()

'Store Configuration
if loadConfig() = false then
	call errorDB(langErrConfig,"")
end if

'Get/Set Cart/Order Session
idOrder = sessionCart()

'Get/Set Customer Session
idCust  = sessionCust()

'Determine Action to be taken
action = lCase(Request.Form("action"))
if len(action) = 0 then
	action = lCase(Request.QueryString("action"))
end if

'Add item to cart
if action = "additem" then
	addItem()
else
	'Check that the session is still active
	if isNull(idOrder) then
		errorMsg = langErrCartEmpty
	else
		'Delete item from cart
		if action = "delitem" then
			delItem()
		end if
		'Recalculate cart totals
		if action = "recalc"  then
			reCalc()
		end if
	end if
end if

'Check for errors after updates
if len(trim(errorMsg)) <> 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(errorMsg)
end if

'Check that the cart still has at least 1 item after updates
if cartQty(idOrder) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrCartEmpty)
end if

'If any updates were made to the order, re-check and re-calculate 
'the Order Discount.
if action = "additem" or action = "delitem" or action = "recalc" then
	orderDisc()
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"-->
<%

'Display the shopping cart
showCart() 

%>
<!--#include file="../UserMods/_INCbottom_.asp"-->
<%

call closeDb()

'*************************************************************************
' Add item to cart
'*************************************************************************
sub addItem()

	'Declare variables local to this subroutine
	dim reqOptSel		'Used when checking for "required" options
	dim arrOptions		'Array - Options from FORM - ID
	dim arrOptionsTXT	'Array - Options from FORM - Text Input
	dim arrOptionsDB	'Array - Options in DB     - ID
	dim arrOptionsDBTXT	'Array - Options in DB     - Description
	
	'Validate Product ID
	IDProduct = Request.Form("idProduct")
	if len(IDProduct) = 0 then
		IDProduct = Request.QueryString("idProduct")
	end if
	if not isNumeric(IDProduct) then
		errorMsg = langErrInvProdID
		exit sub
	end if
	
	'Get Product info from the database
	mySQL = "SELECT description,price,sku,stock,weight,taxExempt " _
	      & "FROM   products " _
	      & "WHERE  idProduct = " & validSQL(IDProduct,"I") & " " _
	      & "AND    active = -1"
	set rsTemp = openRSexecute(mySQL)
	if rstemp.eof then
		errorMsg = langErrInvProdID
		exit sub
	else
		Description	= rstemp("description")
		unitPrice	= rstemp("price")
		SKU			= rstemp("sku")
		stock		= rstemp("stock")
		unitWeight	= rstemp("weight")
		taxExempt	= UCase(trim(rstemp("taxExempt")))
	end if
	call closeRS(rsTemp)
	
	'Validate Quantity
	Quantity = Request.Form("quantity")
	if len(Quantity) = 0 then
		Quantity = 1
	end if
	if not quantityValid(quantity,stock,idProduct) then
		exit sub
	end if
	
	'Check if new qty plus existing qty exceeds max for cart
	if Quantity + cartQty(idOrder) > pMaxCartQty then
		errorMsg = langErrMaxOrdQty
		exit sub
	end if
	
	'Get selected options
	arrOptions    = ""
	arrOptionsTXT = ""
	for each f in Request.Form
		if lCase(left(f,11)) = "optidoption" then
		
			'Check that all "required" options were selected
			if Request.Form("REQ" & mid(f,4)) = "Y" then 
				if (Request.Form("TYP" & mid(f,4)) = "S" _
				and Request.Form(f) = "") _
				or (Request.Form("TYP" & mid(f,4)) = "T" _
				and Request.Form("TXT" & mid(f,4)) = "") then
					errorMsg = langErrReqOpt & "'" & Request.Form("DES" & mid(f,4)) & "'."
					exit sub
				end if
			end if
		
			'Create array of option ID's and any associated text 
			'that may have been entered by the user for that option.
			if   isNumeric(Request.Form(f)) _
			and (Request.Form("TYP" & mid(f,4)) = "S" _
			or  (Request.Form("TYP" & mid(f,4)) = "T" _
			and  Request.Form("TXT" & mid(f,4)) <> "")) then
				'Append delimiter to array string
				if len(arrOptions) > 0 then 
					arrOptions    = arrOptions    & "*,*"
					arrOptionsTXT = arrOptionsTXT & "*,*"
				end if
				'Append values to array string.
				arrOptions = arrOptions & Request.Form(f)
				if len(trim(Request.Form("TXT" & mid(f,4)))) = 0 then
					arrOptionsTXT = arrOptionsTXT & " " 'Prevent empty array
				else
					arrOptionsTXT = arrOptionsTXT & validHTML(Request.Form("DES" & mid(f,4))) & " : " & validHTML(Request.Form("TXT" & mid(f,4)))
				end if
			end if
			
		end if
	next
	arrOptions    = split(arrOptions   ,"*,*")
	arrOptionsTXT = split(arrOptionsTXT,"*,*")

	'Notes : 
	'1. To allow the use of BeginTrans and CommitTrans, the cursor 
	'   location must be on the client (adUseClient).
	'2. To retrieve the @@identity (AutoNumber) value of the inserted
	'   record, the cursor location must be on the server.
	
	'Set CursorLocation of the Connection Object to Client
	connTemp.CursorLocation = adUseClient
	
	'BEGIN Transaction
	connTemp.BeginTrans
	
	'If no cart exists, create new cart and session.
	if isNull(idOrder) then
		set rsTemp = openRSopen("cartHead",adUseServer,adOpenKeySet,adLockOptimistic,adCmdTable,0)
		rsTemp.AddNew
		rsTemp("idCust")      = 0
		rsTemp("orderDate")   = now()
		rsTemp("orderDateInt")= dateInt(now())
		rsTemp("orderStatus") = "U"
		rsTemp("auditInfo")   = Request.ServerVariables("REMOTE_ADDR") & "|" & Request.ServerVariables("REMOTE_USER")
		rsTemp.Update
		session(storeID & "idOrder") = rsTemp("idOrder") '@@identity
		idOrder						 = rsTemp("idOrder")
		call closeRS(rsTemp)
	end if

	'Check if item is already in the cart
	IDCartRow = 0
	mySQL = "SELECT idCartRow,Quantity " _
	      & "FROM   cartRows " _
	      & "WHERE  idOrder = "   & validSQL(idOrder,"I") & " " _
	      & "AND    idProduct = " & validSQL(idProduct,"I")
	set rsTemp = openRSexecute(mySQL)
	do while not rstemp.eof
	
		'Get current options and create DB option arrays.
		mySQL = "SELECT idOption,optionDescrip " _
			  & "FROM   cartRowsOptions " _
			  & "WHERE  idCartRow = " & rstemp("idCartRow")
		set rsTemp2 = openRSexecute(mySQL)
		arrOptionsDB	= ""
		arrOptionsDBTXT = ""
		do while not rstemp2.eof
			if len(arrOptionsDB) = 0 then
				arrOptionsDB	= rstemp2("idOption")
				arrOptionsDBTXT = rstemp2("optionDescrip")
			else
				arrOptionsDB	= arrOptionsDB    & "*,*" & rstemp2("idOption")
				arrOptionsDBTXT = arrOptionsDBTXT & "*,*" & rstemp2("optionDescrip")
			end if
			rstemp2.movenext
		loop
		call closeRS(rsTemp2)
		arrOptionsDB	= split(arrOptionsDB,"*,*")
		arrOptionsDBTXT = split(arrOptionsDBTXT,"*,*")
		
		'Check if Form option arrays and DB option arrays are a match.
		if UBound(arrOptions) = UBound(arrOptionsDB) then
			for i = 0 to Ubound(arrOptions)
				if checkArrayMatch(arrOptions(i),arrOptionsDB) then
					if len(trim(arrOptionsTXT(i))) > 0 then
						if not checkArrayMatch(arrOptionsTXT(i),arrOptionsDBTXT) then
							exit for			'NO MATCH - Text
						end if
					end if
				else
					exit for					'NO MATCH - ID
				end if
			next
			if UBound(arrOptions) = i-1 then	'MATCHED
				oldQuantity = rstemp("quantity")
				IDCartRow   = rstemp("idCartRow")
				exit do
			end if
		end if
	
		'Get next Row
		rsTemp.movenext
		
	loop
	call closeRS(rsTemp)
	
	'INSERT new row
	if IDCartRow = 0 then
	
		'Check if item qualifies for discount
		call getItemDiscount(idProduct,Quantity,unitPrice)
			
		'INSERT CartRows
		set rsTemp = openRSopen("cartRows",adUseServer,adOpenKeySet,adLockOptimistic,adCmdTable,0)
		rsTemp.AddNew
		rsTemp("idOrder")    = idOrder
		rsTemp("idProduct")  = IDProduct
		rsTemp("sku")        = SKU
		rsTemp("quantity")   = Quantity
		rsTemp("unitPrice")  = unitPrice
		rsTemp("unitWeight") = unitWeight
		rsTemp("description")= Description
		rsTemp("taxExempt")  = taxExempt
		rsTemp("idDiscProd") = idDiscProd
		rsTemp("discAmt")    = discAmt
		rsTemp.Update
		IDCartRow            = rsTemp("idCartRow") 'Return @@identity
		call closeRS(rsTemp)
		
		'INSERT CartRowsOptions
		for f = LBound(arrOptions) to UBound(arrOptions)
		
			'If the user entered any text for an option, we assign 
			'the user's text input to the option description, else 
			'we assign the option description located in the database.
			if len(trim(arrOptionsTXT(f))) > 0 then
				optionDescrip = "'" & left(validSQL(arrOptionsTXT(f),"A"),250) & "'"
			else
				optionDescrip = "optionDescrip"
			end if
			
			'Get Option Price and Percentage
			mySQL="SELECT priceToAdd, percToAdd " _
				& "FROM   options " _
				& "WHERE  idOption = " & validSQL(arrOptions(f),"I")
			set rsTemp = openRSexecute(mySQL)
			if not rsTemp.eof then
				priceToAdd = getOptionPrice(rsTemp("priceToAdd"),rsTemp("percToAdd"),unitPrice)
			else
				priceToAdd = 0
			end if
			call closeRS(rsTemp)

			'Update cartRowsOptions
			mySQL = "INSERT INTO cartRowsOptions (" _
				  & "idOrder,idCartRow,idOption,optionPrice," _
				  & "optionDescrip,optionWeight,taxExempt) " _
			      & "SELECT " & validSQL(idOrder,"I") & "," _
			      &				validSQL(idCartRow,"I") & "," _
			      &				validSQL(arrOptions(f),"I") & "," _
			      &			    validSQL(priceToAdd,"D") & "," _
			      &			    optionDescrip & "," _
			      &			   "weightToAdd," _
			      &			   "taxExempt " _
			      & "FROM  options " _
			      & "WHERE idOption = " & validSQL(arrOptions(f),"I")
			set rsTemp = openRSexecute(mySQL)
			call closeRS(rsTemp)

		next
		
	'UPDATE existing row
	else
	
		'Calculate new quantity
		newQuantity = oldQuantity + Quantity
	
		'Check if item qualifies for discount
		call getItemDiscount(idProduct,newQuantity,unitPrice)
		
		'Adjust Discount ID for the SQL statement
		if isNull(idDiscProd) then
			idDiscProd = "NULL"
		end if
		
		'Validate quantity again
		if not quantityValid(newQuantity,stock,idProduct) then
			connTemp.RollBackTrans
			exit sub
		end if

		'UPDATE cartRows
		mySQL = "UPDATE cartRows " _
			  & "SET    quantity   = " & validSQL(newQuantity,"I") & ", " _
			  & "       discAmt    = " & validSQL(discAmt,"D")     & ", " _
			  & "       idDiscProd = " & validSQL(idDiscProd,"I")  & " " _
			  & "WHERE  idCartRow = "  & validSQL(idCartRow,"I")
		set rsTemp = openRSexecute(mySQL)
		call closeRS(rsTemp)
		
	end if
	
	'END Transaction
	connTemp.CommitTrans
	
	'Set CursorLocation of the Connection Object back to Server
	connTemp.CursorLocation = adUseServer
	
end sub

'*************************************************************************
' Remove item from cart
'*************************************************************************
sub delItem()

	'Get cart row to delete
	IDCartRow = Request.QueryString("idCartRow")

	'CartRow was not specified or invalid
	if len(IDCartRow) = 0 or not isNumeric(IDCartRow) then
		errorMsg = langErrItemDelete
		exit sub
	end if
	
	'Set CursorLocation of the Connection Object to Client
	connTemp.CursorLocation = adUseClient
	
	'BEGIN Transaction
	connTemp.BeginTrans
	
	'Remove from cartRowsOptions
	mySQL = "DELETE FROM cartRowsOptions " _
		  & "WHERE  idCartRow = " & validSQL(idCartRow,"I") & " " _
		  & "AND    idOrder = "   & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)

	'Remove from cartRows
	mySQL = "DELETE FROM cartRows " _
		  & "WHERE  idCartRow = " & validSQL(idCartRow,"I") & " " _
		  & "AND    idOrder = "   & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)
	
	'END Transaction
	connTemp.CommitTrans
	
	'Set CursorLocation of the Connection Object back to Server
	connTemp.CursorLocation = adUseServer
	
end sub

'*************************************************************************
' Update item quantity
' Update item discounts
' Update order discount code
'*************************************************************************
sub reCalc()

	'Check if cart has items
	if cartQty(idOrder) = 0 then
		errorMsg = langErrCartEmpty
		exit sub
	end if
	
	'Check if new qty plus existing qty exceeds max for cart
	for each f in Request.Form
		if lcase(left(f,4)) = "iqty" and isNumeric(Request.Form(f)) then
			newQuantity = newQuantity + CLng(Request.Form(f))
		end if
	next
	if newQuantity > pMaxCartQty then
		errorMsg = langErrMaxOrdQty
		exit sub
	end if
	
	'Set CursorLocation of the Connection Object to Client
	connTemp.CursorLocation = adUseClient

	'BEGIN Transaction
	connTemp.BeginTrans
	
	'Check the cart in order to identify wich rows have new quantity
	mySQL = "SELECT idCartRow,idProduct,quantity,unitPrice " _
		  & "FROM   cartRows " _
		  & "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	do while not rstemp.eof
	
		'Identify which row to update
	    if Request.Form("iQty" & rstemp("idCartRow")) <> rstemp("quantity") then
			IDCartRow	= rstemp("idCartRow")
			IDProduct   = rstemp("idProduct")
			newQuantity = Request.Form("iQty" & rstemp("idCartRow"))
			unitPrice   = rsTemp("unitPrice")

			'Validate Quantity
			if not quantityValid(newQuantity,stock,idProduct) then
				connTemp.RollBackTrans
				exit sub
			end if

			'Check if item qualifies for discount
			call getItemDiscount(idProduct,newQuantity,unitPrice)
			
			'Adjust Discount ID for the SQL statement
			if isNull(idDiscProd) then
				idDiscProd = "NULL"
			end if
			
			'Update cart quantity and discount info
			mySQL = "UPDATE cartRows " _
				  & "SET    quantity   = " & validSQL(newQuantity,"I") & ", " _
				  & "       discAmt    = " & validSQL(discAmt,"D")     & ", " _
				  & "       idDiscProd = " & validSQL(idDiscProd,"I")  & " " _
				  & "WHERE  idCartRow = "  & validSQL(idCartRow,"I")
			set rsTemp2 = openRSexecute(mySQL)
			call closeRS(rsTemp2)
			
		end if
		
		rstemp.movenext
		
	loop
	call closeRS(rsTemp)
	
	'Update the discount code with whatever was entered on the form,
	'and reset the discPerc to null or 0. The validity of the 
	'discount code in relation to this particular order is checked 
	'later via a common routine that is called every time ANY type 
	'of update to the order is made. 
	
	'Get Discount Code from Form
	discCode = validHTML(Request.Form("discCode"))
	
	'Update cartHead
	if len(discCode)=0 or isNull(discCode) then
		call updateOrderDisc(idOrder,"","")
	else
		call updateOrderDisc(idOrder,discCode,0)
	end if
	
	'END Transaction
	connTemp.CommitTrans
	
	'Set CursorLocation of the Connection Object back to Server
	connTemp.CursorLocation = adUseServer
	
end sub

'*************************************************************************
' Validate Discount Code
' Update as required
'*************************************************************************
sub orderDisc()

	'Declare variables local to this subroutine
	dim discDateInt	'Date in internal integer format
	dim discTotal	'Order discount total amount
	dim Total       'Order total minus order discount
	
	'Retrieve discount code from cart header
	mySQL = "SELECT discCode " _
	      & "FROM   cartHead " _
	      & "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	if rsTemp.EOF then
		errorMsgDisc = langErrInvOrder
		exit sub
	else
		if isNull(rsTemp("discCode")) then
			discCode = ""
		else
			discCode = rsTemp("discCode")
		end if
	end if
	call closeRS(rsTemp)
	
	'If no discount code is available, update discount info to 
	'nulls just to be safe, and exit this routine.
	if discCode = "" then
		call updateOrderDisc(idOrder,"","")
		exit sub
	end if

	'Get current date in internal integer format so we can compare 
	'it to the date range on the order discount file.
	discDateInt = "" _
		& year(now()) _
		& left("00",2-len(datePart("m",now()))) & datePart("m",now()) _
		& left("00",2-len(datePart("d",now()))) & datePart("d",now())

	'Check if discount code is valid, and still active
	mySQL="SELECT discCode,discPerc,discAmt,discFromAmt,discToAmt " _
		& "FROM   discOrder " _
		& "WHERE  discCode = '" & validSQL(discCode,"A") & "' " _
		& "AND    discStatus = 'A' " _
		& "AND    discValidFrom <= '" & validSQL(discDateInt,"A") & "' " _
		& "AND    discValidTo   >= '" & validSQL(discDateInt,"A") & "' " _
		& "ORDER  BY idDiscOrder "
	set rsTemp = openRSexecute(mySQL)
	if rsTemp.EOF then
		errorMsgDisc = langErrInvDiscCode
		call updateOrderDisc(idOrder,discCode,0)
		exit sub
	else
		discPerc    = rsTemp("discPerc")
		discAmt		= rsTemp("discAmt")
		discFromAmt = rsTemp("discFromAmt")
		discToAmt   = rsTemp("discToAmt")
	end if
	call closeRS(rsTemp)
	
	'Calculate order total (minus the order discount)
	Total = cartTotalExDisc(idOrder,0)
	
	'Compare order total to order total range on order discount file
	if Total < discFromAmt or Total > discToAmt then
		errorMsgDisc = langErrInvDiscAmt1 _
			& pCurrencySign & moneyS(discFromAmt) & " - " _
			& pCurrencySign & moneyS(discToAmt) '& langErrInvDiscAmt2
		call updateOrderDisc(idOrder,discCode,0)
		exit sub
	end if
	
	'If the order discount is NOT based on a percentage, but a fixed 
	'amount, calculate the fixed amount as a percentage of the order.
	if not isNull(discAmt) then
		discPerc = (discAmt / Total) * 100
	end if
	
	'Just in case the percentages are out of bounds after calculations
	if discPerc < 0 then
		discPerc = 0
	end if
	if discPerc > 100 then
		discPerc = 100
	end if
	
	'If we made it this far everything is OK, so we update the cart 
	'header with the discount percentage for the discount code. Note 
	'that the order discount total (discTotal) is not updated here, 
	'but later during the checkout process along with all the other 
	'totals.
	call updateOrderDisc(idOrder,discCode,discPerc)
		
end sub

'*************************************************************************
' Display the contents of the shopping cart
'*************************************************************************
sub showCart()

	'Declare variables local to this subroutine
	dim discTotal	'Order discount amount
	dim optTotal	'Total for item's options (per item)
	dim itemTotal	'Total per item including options and item discounts
	dim Total		'Total for order
%>
	<table width="100%" border="0" cellspacing="0" cellpadding="2">
		<tr><td valign=middle class="CPpageHead">
			<b><%=langGenShoppingCart%></b><br>
		</td></tr>
	</table>
	
	<img src="../UserMods/misc_cleardot.gif" height=4 width=1><br>

	<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<form method="post" name="recalculate" action="cart.asp">
	<input type="hidden" name="action" value="recalc">
	<tr> 
		<td class="CPgenHeadings" nowrap width="5%" ><b><%=langGenQty%></b></td>
		<td class="CPgenHeadings" nowrap width="5%" ><b><%=langGenSKU%></b></td>
		<td class="CPgenHeadings" nowrap width="80%"><b><%=langGenItemDesc%></b></td>
		<td class="CPgenHeadings" nowrap width="5%" ><b><%=langGenSubTotal%></b></td>
		<td class="CPgenHeadings" nowrap width="5%" >&nbsp;</td>
	</tr>
	<tr>
		<td colspan=5><img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=1></td>
	</tr>
<%
	'Get discount code and percentage
	mySQL = "SELECT discCode,discPerc " _
	      & "FROM   cartHead " _
	      & "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rsTemp = openRSexecute(mySQL)
	discCode   = rsTemp("discCode")
	discPerc   = rsTemp("discPerc")
	if isNull(discCode) then 
		discCode  = ""   
	end if
	if isNull(discPerc) then
		discPerc  = 0.00
	end if
	call closeRS(rsTemp)

	'Get all rows for this cart
	mySQL = "SELECT idCartRow,idProduct,quantity," _
	      & "       unitPrice,description,sku,discAmt " _
	      & "FROM   cartRows " _
	      & "WHERE  idOrder = " & validSQL(idOrder,"I") & " " _
	      & "ORDER BY description"
	set rsTemp = openRSexecute(mySQL)
			
	do while not rstemp.eof
				
		'Assign record values to local values
		IDCartRow	= rstemp("idCartRow")
		IDProduct	= rstemp("idProduct")
		Quantity	= rstemp("quantity")
		unitPrice	= rstemp("unitPrice")
		Description	= rstemp("description")
		SKU			= rstemp("sku")
		discAmt		= rstemp("discAmt")
		if isNull(discAmt) then
			discAmt = 0.00
		end if
%> 
		<tr>
			<td nowrap valign=top>
				<input type="text" name="iQty<%=IDCartRow%>" size="2" value="<%=Quantity%>"> 
			</td>
			<td nowrap valign=top>
				<a href="prodView.asp?idproduct=<%=IDProduct%>"><%=SKU%></a>&nbsp;<br>
			</td>
			<td valign=top>
<%
				'Write cartRow line (main item)
				Response.Write description & " - <i>" & pCurrencySign & moneyS(unitPrice) & "</i><br>"
				
				'Write Discount (if any)
				if discAmt > 0 then
					Response.Write "* <i>" & langGenDiscount & " - " & pCurrencySign & moneyS(discAmt) & "</i><br>"
				end if

				'Get all options for this row
				optTotal = 0
				mySQL = "SELECT optionDescrip, optionPrice " _
				      & "FROM   cartRowsOptions " _
				      & "WHERE  idCartRow = " & validSQL(idCartRow,"I")
				set rsTemp2 = openRSexecute(mySQL)
				do while not rstemp2.eof
				
					'Assign record values to local values
					optionDescrip = rstemp2("optionDescrip")
					optionPrice	  = rstemp2("optionPrice")
					
					'Write cartRowOptions line(s) (options)
					Response.Write "* <i>" & optionDescrip
					if optionPrice <> 0 then
						Response.Write " - " & pCurrencySign & moneyS(optionPrice)
					end if
					Response.Write "</i><br>"
					
					'Calculate options Sub Total
					optTotal = optTotal + optionPrice        
					
					rstemp2.movenext
				loop
				call closeRS(rsTemp2)
%>
				<!--<img src="../UserMods/misc_cleardot.gif" border=0 width=350 height=1><br>-->
			</td>
			<td nowrap valign=top>
<%
				'Display item total
				itemTotal = Quantity * (optTotal + unitPrice - discAmt)
				Response.Write pCurrencySign & moneyS(itemTotal) & "<br>"
				
				'Add item total to order total
				total = total + itemTotal
%>
			</td>
			<td align=center valign=top>
				<a href="cart.asp?action=delItem&idCartRow=<%=IDCartRow%>"><img src="../UserMods/butt_delete.gif" border="0" hspace="5"></a>
			</td>
		</tr>
<%
		rstemp.moveNext
		if not rsTemp.eof then
%>
			<tr>
				<td colspan="5" valign="middle">
					<img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=4><br>
					<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr><td class="CPlines">
							<img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=1><br>
						</td></tr>
					</table>
					<img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=4><br>
				</td>
			</tr>
<%
		end if
				
	loop
	call closeRS(rsTemp)
%>
	<tr>
		<td colspan=5><img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=2></td>
	</tr>
	<tr>
		<td class="CPgenHeadings" colspan=2 nowrap>
			<b><%=langGenDiscCode%> :</b>
		</td>
		<td class="CPgenHeadings" nowrap align=right>
			<b><%=langGenSubTotal%> : </b>
		</td>
		<td class="CPgenHeadings" nowrap>
<%
			'Display sub total of all items
			Response.Write "<b>" & pCurrencySign & moneyS(total) & "</b>"
%>
		</td>
		<td class="CPgenHeadings">&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2 nowrap>
			<input type="text" name="discCode" size="10" maxlength="20" value="<%=discCode%>" style="FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt;">
		</td>
		<td nowrap align=right>
<%
			if len(trim(errorMsgDisc)) > 0 then
				Response.Write "<i>" & errorMsgDisc & "</i>&nbsp;&nbsp;-&nbsp;&nbsp;"
			end if
			if discPerc > 0 then
				Response.Write "<i>" & formatNumber(discPerc,2) & "%</i>&nbsp;&nbsp;-&nbsp;&nbsp;"
			end if
			Response.Write langGenDiscCode & " : " 
%>
		</td>
		<td nowrap>
<%
			'Calculate order discount
			discTotal = Round(((total * discPerc) / 100),2)
			
			'Display order discount
			Response.Write pCurrencySign & moneyS(discTotal)
			if discTotal > 0 then
				Response.Write "&nbsp;&nbsp;(-)"
			end if
			
			'Subtract order discount from order total.
			total = total - discTotal
%>
		</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td class="CPgenHeadings" colspan=3 nowrap align=right>
			<b><%=langGenTotal%> : </b>
		</td>
		<td class="CPgenHeadings" nowrap>
<%
			'Display order total
			Response.Write "<b>" & pCurrencySign & moneyS(total) & "</b>"
%>
		</td>
		<td class="CPgenHeadings">&nbsp;</td>
	</tr>
	</table>
	
	<br>
	
	<input alt="Update" src="../UserMods/butt_update.gif" type="image" name="Submit" border="0">
	&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="05_Gateway.asp?action=save"><img src="../UserMods/butt_save.gif" border="0"></a>
	&nbsp;&nbsp;&nbsp;&nbsp;
<%
	'Continue Shopping Button
	idProduct = Request("idProduct")
	if len(idProduct) > 0 and isNumeric(idProduct) then
		Response.Write "<a href=""prodView.asp?idProduct=" & idProduct & """><img src=""../UserMods/butt_shop.gif"" border=0></a>"
	else
		Response.Write "<a href=""prodList.asp""><img src=""../UserMods/butt_shop.gif"" border=0></a>"
	end if
%>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="05_Gateway.asp?action=checkout"><img src="../UserMods/butt_checkout.gif" border="0"></a>
	
	</form>
<%
end sub

'*************************************************************************
'Scan Array for possible match
'*************************************************************************
function checkArrayMatch(tempStr, array1)
	dim i
	checkArrayMatch = false
	tempStr = Lcase(CStr(tempStr))
	for i = 0 to Ubound(array1)
		if LCase(CStr(array1(i))) = tempStr then
			checkArrayMatch = true
			exit for
		end if
	next
end function
'*************************************************************************
'Get item's discount ID and amount.
'Assign the ID and amount to variables with page level scope so 
'that they can be used outside the function.
'*************************************************************************
function getItemDiscount(idProduct,itemQty,itemPrice)

	dim rsTemp
	
	'Initialize External variables
	idDiscProd = null
	discAmt    = 0.00
	
	'Check Parameters
	if not isNumeric(idProduct) _
	or not isNumeric(itemQty) _
	or not isNumeric(itemPrice) then
		exit function
	end if

	'Check database for possible discount
	mySQL = "SELECT idDiscProd,discAmt,discPerc " _
	      & "FROM   DiscProd " _
	      & "WHERE  idProduct = " & validSQL(idProduct,"I") & " " _
	      & "AND    " & validSQL(itemQty,"D") & " >= discFromQty " _
	      & "AND    " & validSQL(itemQty,"D") & " <= discToQty "
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.EOF then
		idDiscProd = rsTemp("idDiscProd")
		'If the product discount is a fixed amount, we simply apply
		'the amount, otherwise we calculate the discount based on a 
		'percentage and move the result to the discount amount field.
		if isNull(rsTemp("discPerc")) then
			discAmt = rsTemp("discAmt")
		else
			discAmt = Round(((itemPrice * rsTemp("discPerc")) / 100),2)
		end if
	end if
	call closeRS(rsTemp)

end function
'*************************************************************************
'Update order discount information on cartHead
'Note : Order discount total (discTotal) is updated later along with
'     : all the other order totals.
'*************************************************************************
function updateOrderDisc(idOrder,discCode,discPerc)

	dim rsTemp
	
	'Check Order ID
	if len(idOrder)=0 or not isNumeric(idOrder) then
		exit function
	end if
	
	'Check parameters and update accordingly
	if (len(discCode) = 0 or isNull(discCode)) _
	or (len(discPerc) = 0 or not isNumeric(discPerc)) then
		mySQL = "UPDATE cartHead " _
			  & "SET    discCode  = null, " _
			  & "       discPerc  = null, " _
			  & "       discTotal = null  " _
			  & "WHERE  idOrder = " & validSQL(idOrder,"I")
	else
		mySQL = "UPDATE cartHead " _
			  & "SET    discCode  = '" & validSQL(discCode,"A") & "', " _
			  & "       discPerc  = "  & validSQL(discPerc,"D") & ",  " _
			  & "       discTotal = null " _
			  & "WHERE  idOrder = " & validSQL(idOrder,"I")
	end if
	
	'Update Order Discount info on cartHead
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)
	
end function
'*************************************************************************
'Validate item quantity
'*************************************************************************
function quantityValid(quantity,stock,idProduct)

	dim rsTemp

	'Initialize
	quantityValid = false

	'Check for numeric
	if not IsNumeric(Quantity) then
		errorMsg = langErrInvQty
		exit function
	end if
	
	'Check > 0
	if CLng(Quantity) <= 0 then
		errorMsg = langErrInvQty
		exit function
	end if
	
	'Check max quantity per product
	if CLng(Quantity) > pMaxItemQty then
		errorMsg = langErrMaxItemQty & pMaxItemQty & "."
		exit function
	end if
	
	'Check quantity against available stock if stock level checking 
	'is enabled.
	if pHideAddStockLevel <> -1 then
		if isNumeric(stock) and not(isEmpty(stock) or isNull(stock)) then
			if CLng(Quantity) > CLng(Stock) then
				errorMsg = langErrNoStock
				exit function
			end if
		else
			if isNumeric(idProduct) and not(isEmpty(idProduct) or isNull(idProduct)) then
				mySQL = "SELECT stock " _
				      & "FROM   products " _
				      & "WHERE  idProduct = " & validSQL(idProduct,"I")
				set rsTemp = openRSexecute(mySQL)
				if CLng(Quantity) > CLng(rsTemp("stock")) then
					errorMsg = langErrNoStock
					exit function
				end if
				call closeRS(rsTemp)
			end if
		end if
	end if
	
	'Return
	quantityValid = true

end function
%>