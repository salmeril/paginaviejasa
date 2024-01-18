<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Product Review Maintenance
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

'Reviews
dim idReview
dim revStatus
dim revRating
dim revName
dim revLocation
dim revEmail
dim revSubj
dim revDetail

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
if action <> "edit" and action <> "del" and action <> "bulkdel" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idReview
if action = "edit" or action = "del" then

	idReview = trim(Request.Form("idReview"))
	if len(idReview) = 0 then
		idReview = trim(Request.QueryString("idReview"))
	end if
	if idReview = "" or not isNumeric(idReview) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Review ID.")
	else
		idReview = CLng(idReview)
	end if
	
end if

if action = "edit" then

	'Get Rating
	revRating = trim(Request.Form("revRating"))
	if revRating = "" or not isNumeric(revRating) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Rating.")
	else
		revRating = CLng(revRating)
	end if
	
	'Get Status
	revStatus = trim(Request.Form("revStatus"))
	if revStatus <> "A" and revStatus <> "I" and revStatus <> "R" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Status.")
	end if
	
	'Get Name
	revName = trim(Request.Form("revName"))
	if len(revName) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Name.")
	end if
	
	'Get Location
	revLocation = trim(Request.Form("revLocation"))
	if len(revLocation) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Location.")
	end if
	
	'Get Email
	revEmail = trim(Request.Form("revEmail"))
	if len(revEmail) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Email.")
	end if

	'Get Subject
	revSubj = trim(Request.Form("revSubj"))
	if len(revSubj) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Subject.")
	end if

	'Get detail
	revDetail = trim(Request.Form("revDetail"))
	if len(revDetail) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Review Detail.")
	end if

end if

'EDIT
if action = "edit" then

	'Update Review
	mySQL = "UPDATE reviews SET " _
	      & "revStatus='"	& revStatus						& "'," _
	      & "revRating="	& revRating						& "," _
	      & "revName='"		& replace(revName,"'","''")		& "'," _
	      & "revLocation='"	& replace(revLocation,"'","''")	& "'," _
	      & "revEmail='"	& revEmail						& "'," _
	      & "revSubj='"		& replace(revSubj,"'","''")		& "'," _
	      & "revDetail='"	& replace(revDetail,"'","''")	& "' " _
	      & "WHERE idReview = " & idReview
	set rs = openRSexecute(mySQL)
	
	call closedb()
	Response.Redirect "SA_rev.asp?recallCookie=1&msg=" & server.URLEncode("Review record was Updated.")

end if

'DELETE or BULK DELETE
if action = "del" or action = "bulkdel" then

	'Declare additional variables
	dim delI		'Array index
	dim delArray	'List of records to be deleted

	'Get record (or records) that must be deleted.
	if action = "del" then
		delArray = split(idReview)
	else
		delArray = split(Request.Form("idReview"),",")
	end if

	'Loop through list of records and delete one by one
	for delI = LBound(delArray) to UBound(delArray)
	
		'Delete records
		mySQL = "DELETE FROM reviews " _
		      & "WHERE idReview = " & trim(delArray(delI))
		set rs = openRSexecute(mySQL)
	
	next

	call closedb()
	Response.Redirect "SA_rev.asp?recallCookie=1&msg=" & server.URLEncode("Review(s) were Deleted.")

end if

'Just in case we ever get this far...
call closedb()
Response.Redirect "SA_rev.asp?recallCookie=1"

%>
