<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Options Groups Maintenance
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

'OptionsGroups
dim idOptionGroup
dim optionGroupDesc
dim optionReq
dim optionType

'Options
dim idOption

'optionsXref
dim idOptOptGroup

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
if  action <> "edit"   _ 
and action <> "del"    _
and action <> "add"    _
and action <> "delopt" _
and action <> "addopt" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idOptionGroup
if action = "edit"   _
or action = "del"    _
or action = "delopt" _
or action = "addopt" then
	idOptionGroup = trim(Request.Form("idOptionGroup"))
	if len(idOptionGroup) = 0 then
		idOptionGroup = trim(Request.QueryString("idOptionGroup"))
	end if
	if idOptionGroup = "" or not isNumeric(idOptionGroup) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Group ID.")
	else
		idOptionGroup = CLng(idOptionGroup)
	end if
end if

'Get idOptOptGroup
if action = "delopt" then
	idOptOptGroup = trim(Request.QueryString("recID"))
	if idOptOptGroup = "" or not isNumeric(idOptOptGroup) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid OptionsXref ID.")
	else
		idOptOptGroup = CLng(idOptOptGroup)
	end if
end if

'Get idOption
if action = "addopt" then
	idOption = trim(Request.Form("idOption"))
	if idOption = "" or not isNumeric(idOption) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option ID.")
	else
		idOption = CLng(idOption)
	end if
end if

'Get Option Group Form Variables
if action = "edit" or action = "add" then
	
	'Option Group Description
	optionGroupDesc = trim(Request.Form("optionGroupDesc"))
	optionGroupDesc = replace(optionGroupDesc,"""","") 'To prevent HTML field terminations
	if len(optionGroupDesc) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Group Description.")
	end if
	
	'Type
	optionType = trim(Request.Form("optionType"))
	if optionType <> "S" and optionType <> "T" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Group Type.")
	end if
	
	'Required?
	optionReq = trim(Request.Form("optionReq"))
	if optionReq <> "Y" and optionReq <> "N" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Option Group Required indicator.")
	end if

end if

'Check that Text Input type not linked to more than one option.
If action = "edit" and optionType = "T" then
	mySQL = "SELECT COUNT(*) AS optionCount " _
	      & "FROM   OptionsXref " _
	      & "WHERE  idOptionGroup = " & idOptionGroup
	set rs = openRSexecute(mySQL)
	if rs("optionCount") > 1 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("This Option Group Type can not be linked to more than one Option.")
	end if
	call closeRS(rs)
end if
If action = "addopt" then
	mySQL = "SELECT COUNT(*) AS optionCount " _
	      & "FROM   OptionsXref a " _
	      & "INNER JOIN OptionsGroups b " _
	      & "ON     b.idOptionGroup = a.idOptionGroup " _
	      & "WHERE  a.idOptionGroup = " & idOptionGroup & " " _
	      & "AND    b.optionType = 'T'"
	set rs = openRSexecute(mySQL)
	if rs("optionCount") >= 1 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("This Option Group Type can not be linked to more than one Option.")
	end if
	call closeRS(rs)
end if

'ADD
if action = "add" then

	'Add Record
	mySQL = "INSERT INTO OptionsGroups (" _
	      & "optionGroupDesc,optionReq,optionType" _
	      & ") VALUES ('" _
	      & replace(optionGroupDesc,"'","''")	& "','" _
	      & optionReq							& "','" _
	      & optionType							& "'" _
	      & ")"
	set rs = openRSexecute(mySQL)
	
	'Get idOptionGroup of INSERTed Record
	mySQL = "SELECT MAX(idOptionGroup) AS maxIdOptionGroup " _
		  & "FROM   OptionsGroups "
	set rs = openRSexecute(mySQL)
	idOptionGroup = rs("maxIdOptionGroup")
	call closeRS(rs)

	call closedb()
	Response.Redirect "SA_optGrp_edit.asp?action=edit&recID=" & idOptionGroup & "&msg=" & server.URLEncode("Option Group was added.")
	
end if

'DELETE
if action = "del" then

	'Check Option Group does not have linked Options
	mySQL = "SELECT idOptionGroup " _
	      & "FROM   optionsXref " _
	      & "WHERE  idOptionGroup = " & idOptionGroup
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Option Group can not be deleted if it has Options linked to it.<br><br>Options must be deleted or un-linked first.")
	end if
	call closeRS(rs)
	
	'Check Option Group does not have linked products
	mySQL = "SELECT idOptionGroup " _
	      & "FROM   optionsGroupsXref " _
	      & "WHERE  idOptionGroup = " & idOptionGroup
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Option Group can not be deleted if it is linked to any Products.<br><br>Product(s) must be Deleted or un-linked from the Option Group first.")
	end if
	call closeRS(rs)

	'Delete Record
	mySQL = "DELETE FROM OptionsGroups " _
	      & "WHERE  idOptionGroup = " & idOptionGroup
	set rs = openRSexecute(mySQL)
	call closedb()
	Response.Redirect "SA_optGrp.asp?msg=" & server.URLEncode("Option Group was deleted.")

end if

'EDIT
if action = "edit" then

	'Update Record
	mySQL = "UPDATE OptionsGroups SET " _
	      & "       optionGroupDesc = '" & replace(optionGroupDesc,"'","''") & "'," _
	      & "       optionReq  = '" & optionReq  & "', " _
	      & "       optionType = '" & optionType & "' " _
	      & "WHERE  idOptionGroup = " & idOptionGroup
	set rs = openRSexecute(mySQL)
	call closedb()
	Response.Redirect "SA_optGrp_edit.asp?action=edit&recID=" & idOptionGroup & "&msg=" & server.URLEncode("Option Group was edited.")
	
end if

'DELOPT
if action = "delopt" then

	'Delete records from optionsXref
	mySQL = "DELETE FROM optionsXref " _
	      & "WHERE idOptOptGroup = " & idOptOptGroup
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_optGrp_edit.asp?action=edit&recID=" & idOptionGroup & "&msg=" & server.URLEncode("Option was removed from Option Group.")

end if

'ADDOPT
if action = "addopt" then

	'Add Option to Option Group
	mySQL = "INSERT INTO optionsXref (" _
	      & "idOptionGroup,idOption" _
	      & ") VALUES (" _
	      & idOptionGroup & "," & idOption _
	      & ")"
	set rs = openRSexecute(mySQL)

	call closedb()
	Response.Redirect "SA_optGrp_edit.asp?action=edit&recID=" & idOptionGroup & "&msg=" & server.URLEncode("Option was added to Option Group.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_optGrp.asp"

%>
