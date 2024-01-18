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

'Get action
action = trim(lCase(Request.Form("action")))
if len(action) = 0 then
	action = trim(lCase(Request.QueryString("action")))
end if
if  action <> "edit" _
and action <> "del" _
and action <> "add" _
and action <> "bulkdel" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idOption
if action = "edit" or action = "del" then
	idOption = trim(Request.Form("idOption"))
	if len(idOption) = 0 then
		idOption = trim(Request.QueryString("idOption"))
	end if
	if idOption = "" or not isNumeric(idOption) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option ID.")
	else
		idOption = CLng(idOption)
	end if
end if

if action = "edit" or action = "add" then

	'Get Option Description
	optionDescrip = trim(Request.Form("optionDescrip"))
	optionDescrip = replace(optionDescrip,"""","") 'To prevent HTML field terminations
	if len(optionDescrip) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Description.")
	end if

	'Get Option Price
	priceToAdd = trim(Request.Form("priceToAdd"))
	if priceToAdd = "" then
		pricetoAdd = 0
	end if
	if not isNumeric(priceToAdd) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Price.")
	else
		priceToAdd = Cdbl(priceToAdd)
	end if
	
	'Get Option Percentage
	percToAdd = trim(Request.Form("percToAdd"))
	if percToAdd = "" then
		percToAdd = 0
	end if
	if not isNumeric(percToAdd) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Percentage.")
	else
		percToAdd = Cdbl(percToAdd)
	end if

	'Get Option Weight
	weightToAdd = trim(Request.Form("weightToAdd"))
	if weightToAdd = "" then
		weightToAdd = 0
	end if
	if not isNumeric(weightToAdd) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Weight.")
	else
		weightToAdd = Cdbl(weightToAdd)
	end if

	'Get Tax Exempt indicator
	taxExempt = trim(Request.Form("taxExempt"))
	if taxExempt <> "Y" and taxExempt <> "N" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Tax Exempt Indicator.")
	end if
	
end if

'ADD
if action = "add" then

	'Add Option
	mySQL = "INSERT INTO Options (" _
	      & "optionDescrip,priceToAdd,weightToAdd,taxExempt,percToAdd" _
	      & ") VALUES (" _
	      & "'" & replace(optionDescrip,"'","''")	& "'," _
	      &       priceToAdd						& "," _
	      &       weightToAdd						& "," _
	      & "'" & taxExempt							& "'," _
	      &       percToAdd							& " " _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	'Get idOption of INSERTed Record
	mySQL = "SELECT MAX(idOption) AS maxIdOption " _
		  & "FROM   Options "
	set rs = openRSexecute(mySQL)
	idOption = rs("maxIdOption")
	call closeRS(rs)
	
	call closedb()
	Response.Redirect "SA_opt_edit.asp?action=edit&recID=" & idOption & "&msg=" & server.URLEncode("Option was Added.")
	
end if

'DELETE or BULK DELETE
if action = "del" or action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of idOptions that will be deleted

	'If just one delete is being performed, we populate just the 
	'first position in the delete array, else we populate the array
	'with a list of all the records that were selected for deletion.
	if action = "del" then
		delArray = split(idOption)
	else
		delArray = split(Request.Form("idOption"),",")
	end if

	'Set CursorLocation of the Connection Object to Client
	cn.CursorLocation = adUseClient
	
	'Loop through list of records and delete one by one
	for delI = LBound(delArray) to UBound(delArray)
	
		'BEGIN Transaction
		cn.BeginTrans
	
		'Delete records from optionsXref
		mySQL = "DELETE FROM optionsXref " _
		      & "WHERE idOption = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)
		
		'Delete records from OptionsProdEx
		mySQL = "DELETE FROM OptionsProdEx " _
		      & "WHERE idOption = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'Delete records from Options
		mySQL = "DELETE FROM Options " _
		      & "WHERE idOption = " &  trim(delArray(delI))
		set rs = openRSexecute(mySQL)

		'END Transaction
		cn.CommitTrans
		
	next

	call closedb()
	Response.Redirect "SA_opt.asp?msg=" & server.URLEncode("Option(s) were Deleted.")

end if

'EDIT
if action = "edit" then

	'Update Record
	mySQL = "UPDATE Options SET " _
	      & "       optionDescrip = '" & replace(optionDescrip,"'","''") & "', " _
	      & "       priceToAdd    = "  & priceToAdd  & ","  _
	      & "       weightToAdd   = "  & weightToAdd & ","  _
	      & "       taxExempt     = '" & taxExempt   & "'," _
		  & "       percToAdd     = "  & percToAdd   & " "  _
	      & "WHERE  idOption = " & idOption
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_opt_edit.asp?action=edit&recID=" & idOption & "&msg=" & server.URLEncode("Option was Updated.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_opt.asp"

%>
