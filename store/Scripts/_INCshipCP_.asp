<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Canada Post online shipping rates.
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
dim CPmerchantID
dim CPfromZip
dim CPsizeL
dim CPsizeW
dim CPsizeH

'Parms - Other
dim CPweight
dim CPtoCountry
dim CPtoState
dim CPtoZip

'Work Fields
dim CPshipArray
dim shipArray
dim shipParms
'*************************************************************************

'Assign session arrays to local arrays
shipArray = session(storeID & "shipArray")
shipParms = session(storeID & "shipParms")

'Assign parameter array values to individual variables
CPmerchantID	= shipParms(0)
CPfromZip		= shipParms(1)
CPsizeL			= shipParms(2)
CPsizeW			= shipParms(3)
CPsizeH			= shipParms(4)
CPweight		= shipParms(5)
CPtoCountry		= shipParms(6)
CPtoState		= shipParms(7)
CPtoZip			= shipParms(8)

'Call Canada Post shipping rate function
if CPweight > 0 then
	dim CPi, CPi2
	'Reposition to first available slot on shipArray()
	for CPi2 = 0 to UBound(shipArray)
		if len(shipArray(CPi2,1)) = 0 then
			exit for
		end if
	next
	'Get shipping rates and load shipArray()
	CPshipArray = CPrates()
	if isArray(CPshipArray) then
		for CPi = 0 to UBound(CPshipArray)
			if len(CPshipArray(CPi,1)) > 0 then
				shipArray(CPi2,0) = CPshipArray(CPi,0)
				shipArray(CPi2,1) = CPshipArray(CPi,1)
				CPi2 = CPi2 + 1
			end if
		next
		session(storeID & "shipArray") = shipArray
	else
       'If error was returned and shipArray is empty, show error
		if len(shipArray(0,1)) = 0 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(CPshipArray)
		end if
	end if
end if

'Get shipping rates
function CPrates()

	dim xmlHttp,xmlDom,strXMLSend,strXMLRec
	dim nodesShipment,nodesService,nodesRate,nodesCurrency,nodesError
	dim requestOption
	dim shipRateArr(100,2)
	dim errMsg
	dim i,j

	errMsg = ""

	'Create XML request
	strXMLSend = "" & _
		"<?xml version=""1.0"" ?>" & _
		"<eparcel>" & _
		"<language> en </language>" & _
		"<ratesAndServicesRequest>" & _
		"  <merchantCPCID> " & CPMERCHANTID & " </merchantCPCID>" & _
		"  <lineItems>" & _
		"    <item>" & _
		"      <quantity> 1 </quantity>" & _
		"      <weight> " & CPWEIGHT & " </weight>" & _
		"      <length> " & CPSIZEL  & " </length>" & _
		"      <width> "  & CPSIZEW  & " </width>" & _
		"      <height> " & CPSIZEH  & " </height>" & _
		"      <description> N/A </description>" & _
		"      <readyToShip/>" & _
		"    </item>" & _
		"  </lineItems>" & _
		"  <city> </city>" & _
		"  <provOrState> "	& CPtoState		& " </provOrState>" & _
		"  <country> "		& CPtoCountry	& " </country>" & _
		"  <postalCode> "	& CPtoZip		& " </postalCode>" & _
		"</ratesAndServicesRequest>" & _
		"</eparcel>"
		
	'Send request
	on error resume next
	set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
	if err.number <> 0 then
		err.Clear
		set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP.4.0")
		if err.number <> 0 then
			CPrates = "CP : " & err.Description
			exit function
		end if
	end if
	on error goto 0
	
	'LIVE
	xmlhttp.Open "POST","http://216.191.36.73:30000",false
	
	'TEST
	'xmlhttp.Open "POST","http://206.191.4.228:30000",false
	
	xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	xmlhttp.send strXMLSend
	if xmlhttp.status <> 200 then
		CPrates = "CP : " & xmlhttp.status & " - " & xmlhttp.statusText & "."
		Set xmlhttp = nothing
		exit function
	end if
	strXMLRec = xmlhttp.responsexml.xml
	Set xmlhttp = nothing

	'Process return
	set xmlDom = Server.CreateObject("microsoft.XMLDOM")
	xmlDom.async = "false"
	xmlDom.LoadXML (strXMLRec)
	'Check that this is a valid XML document
	if xmlDom.parseError.errorCode <> 0 then
		errMsg = "CP : Response from Canada Post could not be parsed."
	else
		set nodesError = xmlDom.documentElement.selectSingleNode("error/statusCode")
		if nodesError is nothing then
			j = 0
			set nodesShipment = xmlDom.documentElement.selectNodes("ratesAndServicesResponse/product")
			for each i in nodesShipment
				'Extract XML elements and data
				set nodesService  = i.selectSingleNode("name")
				set nodesRate     = i.selectSingleNode("rate")
				shipRateArr(j,0)  = nodesRate.text
				shipRateArr(j,1)  = "Canada Post - " & nodesService.text
				j = j + 1
			next
			'If no rates returned, then give error.
			if len(shipRateArr(0,0)) = 0 then
				errMsg = "CP : No rates were returned."
			end if
		else
			errMsg = "CP : " & xmlDom.documentElement.selectSingleNode("error/statusMessage").text
		end if
	end if
	set xmlDom = nothing

	if len(errMsg) > 0 then
		CPrates = errMsg
	else
		CPrates = shipRateArr
	end if

end function
%>
