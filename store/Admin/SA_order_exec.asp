<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Order Maintenance
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
<!--#include file="../Scripts/_INCappEmail_.asp"-->
<!--#include file="../Scripts/_INCupdStatus_.asp"-->
<%
'Database
dim mySQL, cn, rs, rs2

'cartHead
dim idOrder
dim idCust
dim orderDate
dim orderDateInt
dim randomKey
dim subTotal
dim taxTotal
dim shipmentTotal
dim Total
dim shipmentMethod
dim name
dim lastName
dim customerCompany
dim phone
dim email
dim address
dim city
dim locState
dim locCountry
dim zip
dim shippingName
dim shippingLastName
dim shippingPhone
dim shippingAddress
dim shippingCity
dim shippingLocState
dim shippingLocCountry
dim shippingZip
dim paymentType
dim cardType
dim cardNumber
dim cardExpMonth
dim cardExpYear
dim cardVerify
dim cardName
dim generalComments
dim orderStatus
dim auditInfo
dim storeComments
dim storeCommentsPriv
dim adjustReason
dim adjustAmount

'cartRows
dim idCartRow
dim idProduct
dim sku
dim quantity
dim unitPrice
dim unitWeight
dim description

'CartRowsOptions
dim idCartRowOption
dim idOption
dim optionPrice
dim optionDescrip

'Work Fields
dim action				'Action to be taken with this order
dim totalNoAdjust		'Order Total without any adjustments applied
dim orderStatusMail		'Email customer when Order Status changes?
dim orderStatusStockAdj	'Adjust stock level when Order Status changes?
dim delUordHours		'Number of hours to delete Unfinalized orders?

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
and action <> "del" _
and action <> "deluord" _
and action <> "bulkdel" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idOrder
if action = "edit" or action = "del" then

	idOrder = trim(Request.Form("idOrder"))
	if len(idOrder) = 0 then
		idOrder = trim(Request.QueryString("idOrder"))
	end if
	if idOrder = "" or not isNumeric(idOrder) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Order Number.")
	else
		idOrder = CLng(idOrder)
	end if
	
end if

'EDIT
if action = "edit" then

	'Get Order Status and validate it
	orderStatus = UCase(trim(Request.Form("orderStatus")))
	if len(orderStatus) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Order Status.")
	end if
	
	'Get Adjustment Amount and Reason, and validate them
	adjustAmount = trim(Request.Form("adjustAmount"))
	adjustReason = trim(replace(Request.Form("AdjustReason"),"""",""))
	if len(adjustAmount) > 0 then 
		if not IsNumeric(adjustAmount) then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Adjustment Amount must be numeric if entered.")
		else
			adjustAmount = CDbl(adjustAmount)
		end if
	else
		adjustAmount = CDbl("0.00")
	end if
	if adjustAmount > 0 and len(adjustReason) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Adjustment Reason required if Adjustment Amount is entered.")
	end if
	
	'Get miscellaneous fields
	Total				= trim(Request.Form("Total"))
	TotalNoAdjust		= trim(Request.Form("TotalNoAdjust"))
	orderStatusMail		= trim(Request.Form("orderStatusMail"))
	orderStatusStockAdj	= trim(Request.Form("orderStatusStockAdj"))
	
	'Check that Total and TotalNoAdjust is valid. This will usually 
	'only be the case for Unfinalized orders that hasn't reached the 
	'checkout phase.
	if (not isNumeric(total)) or (not isNumeric(TotalNoAdjust)) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Order is not complete and can not be edited.")
	end if
	
	'Adjust Order Total with the Adjustment Amount. If Adjustment 
	'Amount was NOT entered, adjustAmount will be 0.00, so the nett 
	'effect would be that no adjustment is made. Similarly, if there 
	'was an adjustment amount previously, and the user now enters
	'0.00 as the new adjustment amount, the Order Total will 
	'effectively be reset to it's original total.
	total = Cdbl(totalNoAdjust) + Cdbl(adjustAmount)
	
	'Get other fields which can be modified
	storeCommentsPriv	= trim(replace(Request.Form("storeCommentsPriv"),"""",""))
	storeComments		= trim(replace(Request.Form("storeComments"),"""",""))
	Email				= trim(replace(Request.Form("Email"),"""",""))
	Name				= trim(replace(Request.Form("name"),"""",""))
	shippingName		= trim(replace(Request.Form("shippingName"),"""",""))
	LastName			= trim(replace(Request.Form("LastName"),"""",""))
	shippingLastName	= trim(replace(Request.Form("shippingLastName"),"""",""))
	Address				= trim(replace(Request.Form("Address"),"""",""))
	shippingAddress		= trim(replace(Request.Form("shippingAddress"),"""",""))
	City				= trim(replace(Request.Form("City"),"""",""))
	ShippingCity		= trim(replace(Request.Form("ShippingCity"),"""",""))
	Zip					= trim(replace(Request.Form("Zip"),"""",""))
	shippingZip			= trim(replace(Request.Form("shippingZip"),"""",""))
	locState			= trim(replace(Request.Form("locState"),"""",""))
	shippingLocState	= trim(replace(Request.Form("shippingLocState"),"""",""))
	locCountry			= trim(replace(Request.Form("locCountry"),"""",""))
	shippingLocCountry	= trim(replace(Request.Form("shippingLocCountry"),"""",""))
	CustomerCompany		= trim(replace(Request.Form("CustomerCompany"),"""",""))
	Phone				= trim(replace(Request.Form("Phone"),"""",""))
	shippingPhone		= trim(replace(Request.Form("shippingPhone"),"""",""))
	shipmentMethod		= trim(replace(Request.Form("shipmentMethod"),"""",""))
	paymentType			= trim(Request.Form("paymentType"))
	
	'Update cartHead
	mySQL="UPDATE cartHead SET " _
		& "storeCommentsPriv='"	& replace(storeCommentsPriv,"'","''")	& "'," _
		& "storeComments='"		& replace(storeComments,"'","''")		& "'," _
		& "Email='"				& replace(Email,"'","''")				& "'," _
		& "Name='"				& replace(name,"'","''")				& "'," _
		& "shippingName='"		& replace(shippingName,"'","''")		& "'," _
		& "LastName='"			& replace(LastName,"'","''")			& "'," _
		& "shippingLastName='"	& replace(shippingLastName,"'","''")	& "'," _
		& "Address='"			& replace(Address,"'","''")				& "'," _
		& "shippingAddress='"	& replace(shippingAddress,"'","''")		& "'," _
		& "City='"				& replace(City,"'","''")				& "'," _
		& "ShippingCity='"		& replace(ShippingCity,"'","''")		& "'," _
		& "Zip='"				& replace(Zip,"'","''")					& "'," _
		& "shippingZip='"		& replace(shippingZip,"'","''")			& "'," _
		& "locState='"			& replace(locState,"'","''")			& "'," _
		& "shippingLocState='"	& replace(shippingLocState,"'","''")	& "'," _
		& "locCountry='"		& replace(locCountry,"'","''")			& "'," _
		& "shippingLocCountry='"& replace(shippingLocCountry,"'","''")	& "'," _
		& "CustomerCompany='"	& replace(CustomerCompany,"'","''")		& "'," _
		& "Phone='"				& replace(Phone,"'","''")				& "'," _
		& "shippingPhone='"		& replace(shippingPhone,"'","''")		& "'," _
		& "shipmentMethod='"	& replace(shipmentMethod,"'","''")		& "'," _
		& "paymentType='"		& replace(paymentType,"'","''")			& "'," _
		& "AdjustReason='"		& replace(adjustReason,"'","''")		& "'," _
		& "AdjustAmount="		& adjustAmount							& "," _
		& "Total="				& total									& " " _
		& "WHERE idOrder = " & idOrder
	set rs = openRSexecute(mySQL)
	
	'Call the Order Status Update routine
	call updOrderStatus(idOrder,orderStatus,orderStatusMail,orderStatusStockAdj,"")
	
	call closedb()
	Response.Redirect "SA_order.asp?recallCookie=1&msg=" & server.URLEncode("Order was Updated.")

end if

'DELETE or BULK DELETE
if action = "del" or action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of idOrders that will be deleted

	'If just one delete is being performed, we populate just the 
	'first position in the delete array, else we populate the array
	'with a list of all the orders that were selected for deletion.
	if action = "del" then
		delArray = split(idOrder)
	else
		delArray = split(Request.Form("idOrder"),",")
	end if
	
	'Set CursorLocation of the Connection Object to Client
	cn.CursorLocation = adUseClient
	
	'Loop through list of orders and delete one by one
	for delI = LBound(delArray) to UBound(delArray)

		'BEGIN Transaction
		cn.BeginTrans
	
		'Delete records from cartHead
		mySQL = "DELETE FROM cartHead " _
		      & "WHERE idOrder = " & trim(delArray(delI))
		set rs = openRSexecute(mySQL)
	
		'Delete records from cartRows
		mySQL = "DELETE FROM cartRows " _
		      & "WHERE idOrder = " & trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'Delete records from cartRowsOptions
		mySQL = "DELETE FROM cartRowsOptions " _
		      & "WHERE idOrder = " & trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'END Transaction
		cn.CommitTrans
		
	next

	call closedb()
	Response.Redirect "SA_order.asp?recallCookie=1&msg=" & server.URLEncode("Selected Order(s) were Deleted.")

end if

'DELETE Unfinalized orders
if action = "deluord" then

	'Deleted order counter
	dim delOrderCount
	delOrderCount = 0

	'Get delUordHours and validate it
	delUordHours = trim(Request.Form("delUordHours"))
	if len(delUordHours) = 0 then
		delUordHours = trim(Request.QueryString("delUordHours"))
	end if
	if delUordHours = "" or not isNumeric(delUordHours) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid number of Hours selected.")
	else
		delUordHours = CLng(delUordHours)
	end if

	'Read cartHead
	mySQL = "SELECT idOrder " _
	      & "FROM   cartHead " _
	      & "WHERE  orderDateInt < '" & dateInt(dateAdd("h",(delUordHours*-1),now())) & "' " _
		  & "AND    orderStatus  = 'U' "
	set rs = openRSexecute(mySQL)
	do while not rs.eof
	
		'Increment counter
		delOrderCount = delOrderCount + 1

		'Get Order ID
		idOrder = rs("idOrder")
		
		'Delete records from cartRowsOptions
		mySQL = "DELETE FROM cartRowsOptions " _
		      & "WHERE idOrder = " & idOrder
		set rs2 = openRSexecute(mySQL)
		
		'Delete records from cartRows
		mySQL = "DELETE FROM cartRows " _
		      & "WHERE idOrder = " & idOrder
		set rs2 = openRSexecute(mySQL)

		'Delete records from cartHead
		mySQL = "DELETE FROM cartHead " _
		      & "WHERE idOrder = " & idOrder
		set rs2 = openRSexecute(mySQL)
		
		rs.movenext
	loop
	call closeRS(rs)

	call closedb()
	Response.Redirect "SA_order.asp?recallCookie=1&msg=" & server.URLEncode(delOrderCount & " Unfinalized order(s) older than " & delUordHours & " hour(s) were Deleted.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_order.asp?recallCookie=1"
%>
