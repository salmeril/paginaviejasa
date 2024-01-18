<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Order Status update functions
'          : These functions use several other functions located 
'          : in some of the other _INC_ files.
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************

'*************************************************************************
'Update the Order's Status, Inventory, Send Email to customer.
'idOrder			 = Order ID
'orderStatus         = NEW Order Status
'orderStatusMail     = Send Email if status changes? (Y/N)
'orderStatusStockAdj = Adjust Stock if status changes? (Y/N)
'privateText         = Text that will be appended to storeCommentsPriv 
'*************************************************************************
function updOrderStatus(idOrder,orderStatus,orderStatusMail,orderStatusStockAdj,privateText)

	'Declare some variables
	dim mySQL, rs, rs2		'Database variables
	dim customerEmail		'Work field holding Email message body
	dim stockAdj			'(-/+) sign. Used when updating stock levels
	dim orderStatusOld		'Current Order Status
	dim storeComments		'cartHead Field
	dim storeCommentsPriv	'cartHead Field
	dim orderDate			'cartHead Field
	dim Total				'cartHead Field
	dim Name				'cartHead Field
	dim LastName			'cartHead Field
	dim Email				'cartHead Field

	'Miscellaneous data manipulation to ensure conformity
	orderStatus			= UCase(orderStatus)
	orderStatusMail     = UCase(orderStatusMail)
	orderStatusStockAdj	= UCase(orderStatusStockAdj)
	
	'Do some basic checks on the parameters
	if not(isNumeric(idOrder) and len(orderStatus)=1) then
		updOrderStatus = langErrInvParms
		exit function
	end if
	
	'Default the Flags if they were not passed
	if orderStatusMail <> "Y" then
		orderStatusMail = "N"
	end if
	if orderStatusStockAdj <> "Y" then
		orderStatusStockAdj = "N"
	end if
	
	'Get the current Order Record
	mySQL	= "SELECT orderDate,Total,name,lastName,email," _
		    & "       orderStatus,storeComments,storeCommentsPriv " _
			& "FROM   cartHead " _
			& "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rs = openRSexecute(mySQL)
	if rs.eof then
		updOrderStatus = langErrInvOrder
		exit function
	else
		storeComments		= trim(rs("storeComments"))
		storeCommentsPriv	= trim(rs("storeCommentsPriv"))
		orderDate			= rs("orderDate")
		Total				= rs("Total")
		Name				= trim(rs("name"))
		LastName			= trim(rs("LastName"))
		Email				= trim(rs("Email"))
		orderStatusOld		= UCase(trim(rs("orderStatus")))
	end if
	call closeRS(rs)
	
	'Check if order status changed.
	if orderStatus = orderStatusOld then
		updOrderStatus = "" 'No error is generated as this is valid
		exit function
	end if
	
	'Append additional comments to storeComments field
	if len(storeComments) > 0 then
		if right(storeComments,1) <> chr(10) then
			storeComments = storeComments & vbCrLf
		end if
	end if
	storeComments = storeComments _
		& formatTheDate(now()) & " " & time() & vbCrLf _
		& langGenOrderStatus & " : " & orderStatusDesc(orderStatus)
	
	'Append additional comments to storeCommentsPriv field
	if len(storeCommentsPriv) > 0 then
		if right(storeCommentsPriv,1) <> chr(10) then
			storeCommentsPriv = storeCommentsPriv & vbCrLf
		end if
	end if
	storeCommentsPriv = storeCommentsPriv & privateText
	
	'Update cartHead
	mySQL="UPDATE cartHead SET " _
		& "storeCommentsPriv = '" & validSQL(storeCommentsPriv,"A") & "'," _
		& "storeComments = '"     & validSQL(storeComments,"A")     & "'," _
		& "orderStatus = '"       & validSQL(orderStatus,"A")       & "' " _
		& "WHERE idOrder = " & validSQL(idOrder,"I")
	set rs = openRSexecute(mySQL)
	
	'Update Stock levels (if required)
	if UCase(orderStatusStockAdj) = "Y" then
	
		'Decide if stock is to be increased or decreased
		if statUpdPending = -1 then
			select case orderStatusOld
			case "U","S","9"
				select case orderStatus
				case "0","1","2","7"
					stockAdj = "-"	'Decrease
				end select
			case "0","1","2","7"
				select case orderStatus
				case "U","S","9"
					stockAdj = "+"	'Increase
				end select
			end select
		else
			select case orderStatusOld
			case "U","S","0","9"
				select case orderStatus
				case "1","2","7"
					stockAdj = "-"	'Decrease
				end select
			case "1","2","7"
				select case orderStatus
				case "U","S","0","9"
					stockAdj = "+"	'Increase
				end select
			end select
		end if

		'Update Stock Levels (if required)
		if len(trim(stockAdj)) > 0 then
			'Read cartRows to obtain Product ID and Quantity
			mySQL="SELECT idProduct, quantity " _
			    & "FROM   cartRows " _
			    & "WHERE  idOrder = " & validSQL(idOrder,"I")
			set rs = openRSexecute(mySQL)
			do while not rs.eof
				'Update products
				mySQL="UPDATE products SET " _
					& "stock = (stock" & stockAdj & rs("quantity") & ") " _
					& "WHERE idProduct = " & rs("idProduct")
				set rs2 = openRSexecute(mySQL)
				rs.movenext
			loop
			call closeRS(rs)
		end if
		
	end if
	
	'Send Email to Customer
	if UCase(orderStatusMail) = "Y" then
	
		'Build Email Body
		customerEmail = ""
		mySQL = "SELECT configValLong " _
			&   "FROM   storeAdmin " _
			&   "WHERE  configVar = 'statusUpdateEmail' " _
			&   "AND    adminType = 'T'"
		set rs = openRSexecute(mySQL)
		if not rs.eof then
			customerEmail = trim(rs("configValLong"))
		end if
		call closeRS(rs)
	
		'Check for tags and replace
		customerEmail = replace(customerEmail,"#NAME#",name & " " & lastname)
		customerEmail = replace(customerEmail,"#STAT#",orderStatusDesc(orderStatus))
		customerEmail = replace(customerEmail,"#ORDER#",pOrderPrefix & "-" & idOrder)
		customerEmail = replace(customerEmail,"#DATE#",formatTheDate(orderDate))
		customerEmail = replace(customerEmail,"#TOTAL#",pCurrencySign & moneyS(Total))
		customerEmail = replace(customerEmail,"#STORE#",pCompany)
		customerEmail = replace(customerEmail,"#SALES#",pEmailSales)

		'Send email
		call sendmail (pCompany, pEmailSales, Email, langGenOrderNumber & " " & pOrderPrefix & "-" & idOrder, customerEmail, 0)
	end if
	
end function

'*************************************************************************
'Update the Order's Private Comments field.
'idOrder	 = Order ID
'privateText = Text that will be appended to storeCommentsPriv 
'*************************************************************************
function updOrderPrivate(idOrder,privateText)

	'Declare some variables
	dim mySQL, rs, rs2		'Database variables
	dim storeCommentsPriv	'cartHead Field

	'Check Parameters
	if not(isNumeric(idOrder) and len(trim(privateText))>0) then
		updOrderPrivate = langErrInvParms
		exit function
	end if
	
	'Get the Order Record
	mySQL	= "SELECT storeCommentsPriv " _
			& "FROM   cartHead " _
			& "WHERE  idOrder = " & validSQL(idOrder,"I")
	set rs = openRSexecute(mySQL)
	if rs.eof then
		updOrderPrivate = langErrInvOrder
		exit function
	else
		storeCommentsPriv = trim(rs("storeCommentsPriv"))
	end if
	call closeRS(rs)
	
	'Append additional comments to storeCommentsPriv field
	if len(storeCommentsPriv) > 0 then
		if right(storeCommentsPriv,1) <> chr(10) then
			storeCommentsPriv = storeCommentsPriv & vbCrLf
		end if
	end if
	storeCommentsPriv = storeCommentsPriv & privateText
	
	'Update cartHead
	mySQL="UPDATE cartHead SET " _
		& "storeCommentsPriv = '" & validSQL(storeCommentsPriv,"A") & "' " _
		& "WHERE idOrder = " & validSQL(idOrder,"I")
	set rs = openRSexecute(mySQL)
	
end function
%>