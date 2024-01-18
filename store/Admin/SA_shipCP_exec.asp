<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Canada Post Online Shipping Rates
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
const adminLevel = 0
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<%
'Form variables
dim CPactive
dim CPmerchantID
dim CPfromZip
dim CPsizeL
dim CPsizeW
dim CPsizeH

'Database variables
dim mySQL, cn, rs

'Work variables
dim testStr
'*************************************************************************

'Are we in demo mode?
if demoMode = "Y" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("DEMO MODE. Sorry, this featured is NOT available in Demo Mode.")
end if

'Get variables from the Form
CPactive		= trim(Request.Form("CPactive"))
CPmerchantID	= trim(Request.Form("CPmerchantID"))
CPfromZip		= trim(Request.Form("CPfromZip"))
CPsizeL			= trim(Request.Form("CPsizeL"))
CPsizeW			= trim(Request.Form("CPsizeW"))
CPsizeH			= trim(Request.Form("CPsizeH"))

'CPactive
if CPactive <> "Y" then
	CPactive = "N"
end if

'CPmerchantID
if len(CPmerchantID) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid User ID.")
end if

'CPfromZip
if len(CPfromZip) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Postal Code.")
end if

'CPsizeL
if len(CPsizeL) = 0 or not(isNumeric(CPsizeL)) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Parcel Length.")
end if

'CPsizeW
if len(CPsizeW) = 0 or not(isNumeric(CPsizeW)) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Parcel Width.")
end if

'CPsizeH
if len(CPsizeH) = 0 or not(isNumeric(CPsizeH)) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Parcel Height.")
end if

'Check that there are no double quotes (") in any of the fields
testStr = CPactive & CPmerchantID & CPfromZip & CPsizeL & CPsizeW & CPsizeH
if instr(testStr,"""") > 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Double Quotes are not allowed in any of the configuration settings.")
end if

'Update the configuration file
call writeConfigFile()

'Back to main script
Response.Redirect "SA_shipCP.asp?msg=" & server.URLEncode("Online Canada Post was Updated.")

'**********************************************************************
'Update the Configurations
'**********************************************************************
sub writeConfigFile()

	'Create an array of variables which will use to drive the update
	dim configArr, I
	configArr = "CPactive,CPmerchantID,CPfromZip,CPsizeL,CPsizeW,CPsizeH"
	configArr = split(configArr,",")
	
	'Open Database
	call openDb()

	'Loop through array and UPDATE / INSERT new config settings
	for I = 0 to UBound(configArr)
		if len(trim(configArr(I))) > 0 then
			mySQL = "SELECT configVal " _
				&   "FROM   storeAdmin " _
				&   "WHERE  configVar='" & configArr(I) & "' " _
				&   "AND    adminType='S'"
			set rs = openRSexecute(mySQL)
			if rs.EOF then
				'INSERT
				call closeRS(rs)
				mySQL = "INSERT INTO storeAdmin " _
					  & "(adminType,configVar,configVal) " _
					  & "VALUES " _
					  & "('S','" & configArr(I) & "','" & replace(eval(configArr(I)),"'","''") & "')"
				set rs = openRSexecute(mySQL)
			else
				if trim(rs("configVal")) = eval(configArr(I)) then
					'IGNORE
					call closeRS(rs)
				else
					'UPDATE
					call closeRS(rs)
					mySQL = "UPDATE storeAdmin SET " _
					      & "       configVal='" & replace(eval(configArr(I)),"'","''") & "' " _
						  & "WHERE  configVar='" & configArr(I) & "' " _
						  & "AND    adminType='S'"
					set rs = openRSexecute(mySQL)
				end if
			end if
		end if
	next

	'Close Database
	call closedb()
	
end sub
%>
