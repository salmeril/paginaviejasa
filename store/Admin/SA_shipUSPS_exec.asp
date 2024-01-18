<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : USPS Online Shipping Rates
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
dim USPSactive
dim USPSUserID
dim USPSPassword
dim USPSfromZip
dim USPSservice
dim USPSintNtl
dim USPSsize
dim USPSmachinable

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
USPSactive		= trim(Request.Form("USPSactive"))
USPSUserID		= trim(Request.Form("USPSUserID"))
USPSPassword	= trim(Request.Form("USPSPassword"))
USPSfromZip		= trim(Request.Form("USPSfromZip"))
USPSservice		= trim(Request.Form("USPSservice1")) & "," _
				& trim(Request.Form("USPSservice2")) & "," _
				& trim(Request.Form("USPSservice3")) & "," _
				& trim(Request.Form("USPSservice4"))
USPSintNtl		= trim(Request.Form("USPSintNtl"))
USPSsize		= trim(Request.Form("USPSsize"))
USPSmachinable	= trim(Request.Form("USPSmachinable"))

'USPSactive
if USPSactive <> "Y" then
	USPSactive = "N"
end if

'USPSUserID
if len(USPSUserID) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid USPS User Name.")
end if

'USPSPassword
if len(USPSPassword) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid USPS Password.")
end if

'USPSfromZip
if len(USPSfromZip) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Zip/Postal Code.")
end if

'USPSservice - Get rid of double commas
do while instr(USPSservice,",,") > 0
	USPSservice = replace(USPSservice,",,",",")
loop

'USPSservice - Get rid of leading and trailing commas
USPSservice = "*" & USPSservice & "*"
USPSservice = replace(USPSservice,"*,","")
USPSservice = replace(USPSservice,",*","")
USPSservice = replace(USPSservice,"*","")
USPSservice = trim(USPSservice)

'USPSservice - Check for valid value
if len(USPSservice) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("At least one Service must be selected.")
end if

'USPSintNtl
if USPSintNtl <> "Y" then
	USPSintNtl = "N"
end if

'Check that there are no double quotes (") in any of the fields
testStr = USPSactive & USPSUserID & USPSPassword & USPSfromZip _
	& USPSservice & USPSintNtl & USPSsize & USPSmachinable
if instr(testStr,"""") > 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Double Quotes are not allowed in any of the configuration settings.")
end if

'Update the configuration file
call writeConfigFile()

'Back to main script
Response.Redirect "SA_shipUSPS.asp?msg=" & server.URLEncode("Online USPS was Updated.")

'**********************************************************************
'Update the Configurations
'**********************************************************************
sub writeConfigFile()

	'Create an array of variables which will use to drive the update
	dim configArr, I
	configArr = "" _
		& "USPSactive,USPSUserID,USPSPassword,USPSfromZip," _
		& "USPSservice,USPSintNtl,USPSsize,USPSmachinable"
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
