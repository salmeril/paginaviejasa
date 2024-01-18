<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Bulk Order Maintenance
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

'cartRows
dim idCartRow
dim idOrder
dim idProduct
dim sku
dim quantity
dim unitPrice
dim unitWeight
dim description
dim downloadCount
dim downloadDate

'Work Fields
dim action
dim idProduct1
dim idProduct2

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
if action <> "additem" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idProduct(s), Price, Quantity
if action = "additem" then

	'First Item (the one that we will be looking for)
	idProduct1 = trim(Request.Form("idProduct1"))
	if len(idProduct1) = 0 then
		idProduct1 = trim(Request.QueryString("idProduct1"))
	end if
	if idProduct1 = "" or not isNumeric(idProduct1) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Item(s) selected.")
	else
		idProduct1 = CLng(idProduct1)
	end if
	
	'Second Item (the one the order will be modified with)
	idProduct2 = trim(Request.Form("idProduct2"))
	if len(idProduct2) = 0 then
		idProduct2 = trim(Request.QueryString("idProduct2"))
	end if
	if idProduct2 = "" or not isNumeric(idProduct2) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Item(s) selected.")
	else
		idProduct2 = CLng(idProduct2)
	end if
	
	'Get some of the Second Item's details for later use
	mySQL="SELECT idProduct,sku,description,weight,price " _
	    & "FROM   products " _
	    & "WHERE  idProduct=" & idProduct2
	set rs = openRSexecute(mySQL)
	if rs.EOF then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("The requested Product could not be found.")
	else
		idProduct	= rs("idProduct")
		sku			= rs("sku")
		description	= rs("description")
		unitWeight	= rs("weight")
		unitPrice	= rs("price")
		Quantity    = 1
		if Request.Form("invertPrice") = "Y" then
			unitPrice = unitPrice * -1
		end if
	end if
	call closeRS(rs)
	
end if

'ADD Item
if action = "additem" then

	'Set CursorLocation of the Connection Object to Client
	cn.CursorLocation = adUseClient
	
	'BEGIN Transaction
	cn.BeginTrans
	
	'INSERT the row(s)
	mysQL = "INSERT INTO cartRows " _
		  & "      (idOrder,idProduct,sku,quantity," _
		  & "       unitPrice,unitWeight,description) " _
		  & "SELECT cartRows.idOrder," _
		  &         idProduct   & ",'" _
		  &         sku         & "'," _
		  &         quantity    & "," _
		  &         unitPrice   & "," _
		  &         unitWeight  & ",'" _
		  &         replace(description,"'","''") & "' " _ 
		  & "FROM   cartRows " _
		  & "WHERE  cartRows.idProduct=" & idProduct1
	set rs = openRSexecute(mySQL)
		  
	'END Transaction
	cn.CommitTrans

	call closedb()
	Response.Redirect "SA_order_bulk.asp?msg=" & server.URLEncode("Items were added to Order(s).")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_order_bulk.asp"
%>
