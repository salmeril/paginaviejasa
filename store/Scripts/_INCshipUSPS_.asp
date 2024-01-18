<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : USPS online shipping rates.
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
<%
'Parms - Database
dim USPSUserID
dim USPSPassword
dim USPSfromZip
dim USPSservice
dim USPSsize
dim USPSmachinable
dim USPScontainer

'Parms - Other
dim USPSWeight
dim USPStoCountry
dim USPStoZip

'Work Fields
dim USPSshipArray
dim shipArray
dim shipParms
'*************************************************************************

'Assign session arrays to local arrays
shipArray = session(storeID & "shipArray")
shipParms = session(storeID & "shipParms")

'Assign parameter array values to individual variables
USPSuserID		= shipParms(0)
USPSpassword	= shipParms(1)
USPSfromZip		= shipParms(2)
USPSservice		= shipParms(3)
USPSsize		= shipParms(4)
USPSmachinable	= shipParms(5)
USPSweight		= shipParms(6)
USPStoCountry	= shipParms(7)
USPStoZip		= shipParms(8)
USPScontainer	= "NONE"

'Cater for weights under 1 pound
if USPSweight > 0 and USPSweight < 1 then
	USPSweight = 1
end if

'Call UPS shipping rate function
if USPSweight > 0 then
	dim USPSi, USPSi2
	'Reposition to first available slot on shipArray()
	for USPSi2 = 0 to UBound(shipArray)
		if len(shipArray(USPSi2,1)) = 0 then
			exit for
		end if
	next
	'Get shipping rates and load shipArray()
	USPSshipArray = USPSrates()
	if isArray(USPSshipArray) then
		for USPSi = 0 to UBound(USPSshipArray)
			if len(USPSshipArray(USPSi,1)) > 0 then
				shipArray(USPSi2,0) = USPSshipArray(USPSi,0)
				shipArray(USPSi2,1) = USPSshipArray(USPSi,1)
				USPSi2 = USPSi2 + 1
			end if
		next
		session(storeID & "shipArray") = shipArray
	else
       'If error was returned and shipArray is empty, show error
		if len(shipArray(0,1)) = 0 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(USPSshipArray)
		end if
	end if
end if

'Get shipping rates
function USPSrates()

	dim xmlHttp,xmlDom,strXMLSend,strXMLRec
	dim nodesShipment,nodesService,nodesRate,nodesError
	dim shipRateArr(100,2), serviceArray
	dim errMsg
	dim i,j

	'Determine which services to get rates for
	serviceArray = split(USPSSERVICE,",")

	'Create XML request
	strXMLSend = "" & _
	"<?xml version=""1.0""?>" & _
	"<RateRequest USERID=""" & USPSUSERID & """ PASSWORD=""" & USPSPASSWORD & """>"
	for i = 0 to UBound(serviceArray)
		strXMLSend = strXMLSend & _
		"<Package ID=""" & i & """>" & _
		"	<Service>" & SERVICEARRAY(i) & "</Service>" & _
		"	<ZipOrigination>" & USPSFROMZIP & "</ZipOrigination>" & _
		"	<ZipDestination>" & USPSTOZIP & "</ZipDestination>" & _
		"	<Pounds>" & Round(USPSWEIGHT,0) & "</Pounds>" & _
		"	<Ounces>0</Ounces>" & _
		"	<Container>" & USPSCONTAINER & "</Container>" & _
		"	<Size>" & USPSSIZE & "</Size>" & _
		"	<Machinable>" & USPSMACHINABLE & "</Machinable>" & _
		"</Package>"
	next
	strXMLSend = strXMLSend & _
	"</RateRequest>"

	'Send request
	on error resume next
	set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
	if err.number <> 0 then
		err.Clear
		set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP.4.0")
		if err.number <> 0 then
			USPSrates = "USPS : " & err.Description
			exit function
		end if
	end if
	on error goto 0
	xmlhttp.Open "POST","http://production.shippingapis.com/ShippingAPI.dll",false
	xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	xmlhttp.send "API=Rate&XML=" & strXMLSend
	if xmlhttp.status <> 200 then
		USPSrates = "USPS : HTTP Error " & xmlhttp.status & " - " & xmlhttp.statusText & "."
		Set xmlhttp = nothing
		exit function
	end if
	strXMLRec = xmlhttp.responseText
	Set xmlhttp = nothing

	'Process return
	set xmlDom = Server.CreateObject("microsoft.XMLDOM")
	xmlDom.async = "false"
	xmlDom.LoadXML (strXMLRec)
	'Check that this is a valid XML document
	if xmlDom.parseError.errorCode <> 0 then
		errMsg = "USPS : Response from USPS could not be parsed."
	else
		'Check for document level error
		set nodesError = xmlDom.documentElement.selectSingleNode("/Error/Description")
		if nodesError is nothing then
			j = 0
			set nodesShipment = xmlDom.documentElement.selectNodes("Package")
			for each i in nodesShipment
				'Extract XML elements and data
				set nodesService  = i.selectSingleNode("Service")
				set nodesRate     = i.selectSingleNode("Postage")
				set nodesError    = i.selectSingleNode("Error/Description")
				if nodesError is nothing then
					shipRateArr(j,0)  = nodesRate.text
					shipRateArr(j,1)  = "USPS - " & nodesService.text
					j = j + 1
				else
					errMsg = "USPS : " & nodesError.Text
				end if
			next
		else
			errMsg = "USPS : " & nodesError.Text
		end if
		'If no rates returned, and no other errors, then give error.
		if len(shipRateArr(0,0)) = 0 then
			if len(errMsg) = 0 then
				errMsg = "USPS : No rates were returned."
			end if
		else
			errMsg = ""
		end if
	end if
	set xmlDom = nothing

	if len(errMsg) > 0 then
		USPSrates = errMsg
	else
		USPSrates = shipRateArr
	end if

end function
%>