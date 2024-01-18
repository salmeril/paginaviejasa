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
and action <> "del" _
and action <> "bulkdel" _
and action <> "add" _
and action <> "root" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idCategory
if action = "edit" or action = "del" then

	idCategory = Request.Form("idCategory")
	if idCategory = "" or not isNumeric(idCategory) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if
	
end if

if action = "edit" or action = "add" then

	'Get Category Description
	categoryDesc = trim(Request.Form("categoryDesc"))
	categoryDesc = replace(categoryDesc,"""","") 'To prevent HTML field terminations
	if len(categoryDesc) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Category Description.")
	end if
	
	'Get idParentCategory
	idParentCategory = Request.Form("idParentCategory")
	if idParentCategory = "" or not isNumeric(idParentCategory) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Parent Category.")
	end if
	
	'Check idParentCategory exists on DB
	'Exclude Root category from this test
	if idParentCategory <> 0 then
		mySQL = "SELECT idCategory " _
		      & "FROM   categories " _
		      & "WHERE  idCategory = " & idParentCategory
		set rs = openRSexecute(mySQL)
		if rs.eof then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Parent Category.")
		end if
		call closeRS(rs)
	end if
	
	'Check idParentCategory not linked to products
	mySQL = "SELECT idCategory " _
	      & "FROM   categories_products " _
	      & "WHERE  idCategory = " & idParentCategory
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("The Parent Category can not have products directly linked to it.")
	end if
	call closeRS(rs)
	
	'Get categoryFeatured
	categoryFeatured = UCase(trim(Request.Form("categoryFeatured")))
	if categoryFeatured <> "Y" and categoryFeatured <> "N" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Featured value.")
	end if
	
	'Get Category HTML
	categoryHTML = trim(Request.Form("categoryHTML"))
	if len(categoryHTML) > 255 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Category HTML can not exceed 255 characters.")
	end if
	
end if

'ADD
if action = "add" then

	'Add Record
	mySQL = "INSERT INTO categories (" _
	      & "categoryDesc, idParentCategory, " _
	      & "categoryFeatured, categoryHTML" _
	      & ") VALUES (" _
	      & "'" & replace(categoryDesc,"'","''") & "'," _
	      &       idParentCategory				 & "," _
	      & "'" & categoryFeatured				 & "'," _
	      & "'" & replace(categoryHTML,"'","''") & "'" _
	      & ")"
	set rs = openRSexecute(mySQL)
	call closedb()
	Response.Redirect "SA_cat.asp?msg=" & server.URLEncode("Category was added.")
	
end if

'DELETE or BULK DELETE
if action = "del" or action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of idCategories that will be deleted
	
	'If just one delete is being performed, we populate just the 
	'first position in the delete array, else we populate the array
	'with a list of all the records that were selected for deletion.
	if action = "del" then
		delArray = split(idCategory)
	else
		delArray = split(Request.Form("idCategory"),",")
	end if
	
	'Set CursorLocation of the Connection Object to Client
	cn.CursorLocation = adUseClient
	
	'Loop through list of records and delete one by one
	for delI = LBound(delArray) to UBound(delArray)
	
		'BEGIN Transaction
		cn.BeginTrans
		
		'Delete Record
		mySQL = "DELETE FROM Categories_Products " _
		      & "WHERE  idCategory = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'Delete Category
		mySQL = "DELETE FROM categories " _
		      & "WHERE  idCategory = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'END Transaction
		cn.CommitTrans
		
	next

	call closedb()
	Response.Redirect "SA_cat.asp?msg=" & server.URLEncode("Category(s) were deleted.")

end if

'EDIT
if action = "edit" then

	'Check idCategory <> idParentCategory
	if idCategory = idParentCategory then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Category can not be linked to itself.")
	end if
	
	'Check that edited Category is not being linked to one of it's 
	'own sub-categories
	call expandCategory(idCategory)

	'Update Record
	mySQL = "UPDATE categories SET " _
	      & "       categoryDesc = '"     & replace(categoryDesc,"'","''") & "'," _
	      & "       idParentCategory = "  & idParentCategory & "," _
	      & "       categoryFeatured = '" & categoryFeatured & "'," _
		  & "       categoryHTML = '"     & replace(categoryHTML,"'","''") & "' " _
	      & "WHERE  idCategory = " & idCategory
	set rs = openRSexecute(mySQL)
	call closedb()
	Response.Redirect "SA_cat.asp?msg=" & server.URLEncode("Category was edited.")
	
end if

'ROOT
if action = "root" then

	'Check no Root Category exists
	mySQL = "SELECT idCategory " _
	      & "FROM   categories " _
	      & "WHERE  idParentCategory = 0"
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Root Category already exists.")
	end if
	call closeRS(rs)

	'Create Root Record
	mySQL = "INSERT INTO categories (" _
	      & "categoryDesc, idParentCategory, " _
	      & "categoryFeatured, categoryHTML" _
	      & ") VALUES (" _
	      & "'Root',0,'N',''" _
	      & ")"
	set rs = openRSexecute(mySQL)
	call closedb()
	Response.Redirect "SA_cat.asp?msg=" & server.URLEncode("Root Category was created.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_cat.asp"

'***********************************************************************
'Check that the Category being edited is not being linked to another 
'Category which is currently acting as one of it's Sub-Categories.
'***********************************************************************
function expandCategory(pIdCategory)

	dim mySQL, rs
	
	mySQL = "SELECT idCategory, idParentCategory " _
		  & "FROM   categories " _
		  & "WHERE  idParentcategory = " & pIdCategory
	set rs = openRSexecute(mySQL)
	do while not rs.eof
		if Clng(idParentCategory) = Clng(rs("idCategory")) then
			call closeDB()
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Category can not be linked to one of it's own Sub-Categories.")
		end if
		call expandCategory(rs("idCategory"))
		rs.movenext
	loop
	call closeRS(rs)
	
end function
%>
