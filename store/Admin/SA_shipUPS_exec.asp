<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : UPS Online Shipping Rates
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
dim UPSactive
dim UPSAccessID
dim UPSUserID
dim UPSPassword
dim UPSfromZip
dim UPSfromCntry
dim UPSpickupType
dim UPSpackType
dim UPSshipCode
dim UPSweightUnit
dim UPSallRates

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
UPSactive		= trim(Request.Form("UPSactive"))
UPSAccessID		= trim(Request.Form("UPSAccessID"))
UPSUserID		= trim(Request.Form("UPSUserID"))
UPSPassword		= trim(Request.Form("UPSPassword"))
UPSfromZip		= trim(Request.Form("UPSfromZip"))
UPSfromCntry	= trim(Request.Form("UPSfromCntry"))
UPSpickupType	= trim(Request.Form("UPSpickupType"))
UPSpackType		= trim(Request.Form("UPSpackType"))
UPSshipCode		= trim(Request.Form("UPSshipCode"))
UPSweightUnit	= trim(Request.Form("UPSweightUnit"))

'UPSAllrates
if len(UPSshipCode) = 0 then
	UPSAllrates = "Y"
else
	UPSAllrates = "N"
end if

'UPSactive
if UPSactive <> "Y" then
	UPSactive = "N"
end if

'UPSAccessID
if len(UPSAccessID) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid UPS Access Key.")
end if

'UPSUserID
if len(UPSUserID) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid UPS User ID.")
end if

'UPSPassword
if len(UPSPassword) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid UPS Password.")
end if

'UPSfromCntry
if len(UPSfromCntry) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Country Code.")
else
	UPSfromCntry = UCase(UPSfromCntry)
end if

'UPSfromZip
if len(UPSfromZip) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Zip/Postal Code.")
end if

'Check that there are no double quotes (") in any of the fields
testStr = UPSactive & UPSAccessID & UPSUserID & UPSPassword _
	& UPSfromZip & UPSfromCntry & UPSpickupType & UPSpackType _
	& UPSshipCode & UPSweightUnit & UPSallRates
if instr(testStr,"""") > 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Double Quotes are not allowed in any of the configuration settings.")
end if

'Update the configuration file
call writeConfigFile()

'Back to main script
Response.Redirect "SA_shipUPS.asp?msg=" & server.URLEncode("Online UPS was Updated.")

'**********************************************************************
'Update the Configurations
'**********************************************************************
sub writeConfigFile()

	'Create an array of variables which will use to drive the update
	dim configArr, I
	configArr = "" _
		& "UPSactive,UPSAccessID,UPSUserID,UPSPassword,UPSfromZip," _
		& "UPSfromCntry,UPSpickupType,UPSpackType,UPSshipCode," _
		& "UPSweightUnit,UPSallRates"
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
