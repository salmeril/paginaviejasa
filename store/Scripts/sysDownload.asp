<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Handles the Downloads for the store
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
%>
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="_INCconfig_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="_INCappFunctions_.asp"-->
<%
'Variables
dim orderStatus
dim randomKey
dim idCartRow
dim fileName
dim filePath
dim qIdOrder
dim downloadCount
dim downloadDate

'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2

'Session
dim idOrder
dim idCust

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	call errorDB(langErrConfig,"")
end if

'Get/Set Cart/Order Session
idOrder = sessionCart()

'Get/Set Customer Session
idCust = sessionCust()

'Get parameters
randomKey  = Request.QueryString("randomKey")
qIdOrder   = Request.QueryString("idOrder")
idCartRow  = Request.QueryString("idCartRow")

'Validate parameters
if len(randomKey) = 0 or (not isNumeric(randomKey)) _
or len(qIdOrder)  = 0 or (not isNumeric(qIdOrder))  _
or len(idCartRow) = 0 or (not isNumeric(idCartRow)) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvParms)
end if

'Check if idOrder/randomKey/idCartRow combination exists
mySQL="SELECT a.orderStatus " _
	& "FROM   cartHead a, cartRows b " _
	& "WHERE  a.randomKey = '" & validSQL(randomKey,"A") & "'" _
	& "AND    a.idOrder   = "  & validSQL(qIdOrder,"I")  & " " _
	& "AND    a.idOrder   = b.idOrder " _
	& "AND    b.idCartRow = "  & validSQL(idCartRow,"I") 
set rsTemp = openRSexecute(mySQL)
if rsTemp.EOF then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvParms)
else
	orderStatus = rsTemp("orderStatus")
end if
call closeRS(rsTemp)

'Check if payment is required
if  orderStatus <> "1" _
and orderStatus <> "2" _
and orderStatus <> "7" then
	if cartTotal(qIdOrder,IdCartRow) > 0 then 'Free Download?
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrDownNotPaid)
	end if
end if

'Get Filename
filename = downloadFile(qIdOrder,IdCartRow)
if len(fileName) = 0 then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvFilename)
end if

'Construct FilePath
filePath = server.MapPath(pDownloadDir)
if right(trim(filePath),1) <> "\" then 'Check for "\" at the end
	filePath = filePath & "\" & fileName
else
	filePath = filePath & fileName
end if

'Get the download counter and first downloaded date
mySQL="SELECT downloadCount, downloadDate " _
	& "FROM   cartRows " _
    & "WHERE  idCartRow = " & validSQL(idCartRow,"I")
set rsTemp = openRSexecute(mySQL)
if not rsTemp.EOF then

	'Get current downloadCount, downloadDate
	downloadCount = rsTemp("downloadCount")
	downloadDate  = rsTemp("downloadDate")
	
	'Adjust downloadCount
	if isNull(downloadCount) or not isNumeric(downloadCount) then
		downloadCount = 1
	else
		downloadCount = downloadCount + 1
	end if

	'Adjust downloadDate (only done on the first download)
	if isNull(downloadDate) or not isNumeric(downloadDate) then
		downloadDate = dateInt(now())
	end if

	'Validate downloadCount not exceeded
	if pMaxDownloadCount <> 0 then
		if downloadCount > pMaxDownloadCount then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrDownMaxTimes)
		end if
	end if
	
	'Validate hours allowed for download is not exceeded
	if pMaxDownloadHours <> 0 then
		if downloadHoursLapsed(downloadDate) > pMaxDownloadHours then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrDownMaxHours)
		end if
	end if
	
	'Update the downloadCounter, downloadDate
	mySQL = "UPDATE cartRows " _
	      & "SET    downloadCount = " & validSQL(downloadCount,"I") & ", " _
	      & "       downloadDate = '" & validSQL(downloadDate,"A")  & "' " _
	      & "WHERE  idCartRow=" & idCartRow
	set rsTemp2 = openRSexecute(mySQL)
	call closeRS(rsTemp2)
	
end if
call closeRS(rsTemp)

'Close the DB
call closeDB()

'If we got this far, everything is OK, so redirect to the file.
Response.Redirect pDownloadDir & fileName

'*********************************************************************
'Calculate difference between the download date and the system date
'*********************************************************************
function downloadHoursLapsed(str1)

	dim tempDate
	
	if len(trim(str1))=14 and isnumeric(str1) then
		tempDate = "" _
			& mid(str1,5,2) & "/" _
			& mid(str1,7,2) & "/" _
			& left(str1,4)  & " " _
			& mid(str1,9,2) & ":" _
			& mid(str1,11,2)
		if IsDate(tempDate) then
			downloadHoursLapsed = dateDiff("h",CDate(tempDate),now())
		else
			downloadHoursLapsed = pMaxDownloadHours + 1 'Force an error
		end if
	else
		downloadHoursLapsed = pMaxDownloadHours + 1 'Force an error
	end if
	
end function
%>