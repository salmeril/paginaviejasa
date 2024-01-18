<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : UPS online shipping rates.
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

'Parms - Other
dim UPSWeight
dim UPStoCountry
dim UPStoZip

'Work Fields
dim UPSshipArray
dim shipArray
dim shipParms
'*************************************************************************

'Assign session arrays to local arrays
shipArray = session(storeID & "shipArray")
shipParms = session(storeID & "shipParms")

'Assign parameter array values to individual variables
UPSaccessID		= shipParms(0)
UPSuserID		= shipParms(1)
UPSpassword		= shipParms(2)
UPSfromZip		= shipParms(3)
UPSfromCntry	= shipParms(4)
UPSpickupType	= shipParms(5)
UPSpackType		= shipParms(6)
UPSshipCode		= shipParms(7)
UPSweightUnit	= shipParms(8)
UPSallRates		= shipParms(9)
UPSweight		= shipParms(10)
UPStoCountry	= shipParms(11)
UPStoZip		= shipParms(12)

'Call UPS shipping rate function
if UPSweight > 0 then
	dim UPSi, UPSi2
	'Reposition to first available slot on shipArray()
	for UPSi2 = 0 to UBound(shipArray)
		if len(shipArray(UPSi2,1)) = 0 then
			exit for
		end if
	next
	'Get shipping rates and load shipArray()
	UPSshipArray = UPSrates()
	if isArray(UPSshipArray) then
		for UPSi = 0 to UBound(UPSshipArray)
			if len(UPSshipArray(UPSi,1)) > 0 then
				shipArray(UPSi2,0) = UPSshipArray(UPSi,0)
				shipArray(UPSi2,1) = UPSshipArray(UPSi,1)
				UPSi2 = UPSi2 + 1
			end if
		next
		session(storeID & "shipArray") = shipArray
	else
       'If error was returned and shipArray is empty, show error
		if len(shipArray(0,1)) = 0 then
			response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(UPSshipArray)
		end if
	end if
end if

'Get shipping rates
function UPSrates()

	dim xmlHttp,xmlDom,strXMLSend,strXMLRec
	dim nodesShipment,nodesService,nodesRate,nodesCurrency,nodesError
	dim requestOption
	dim shipRateArr(100,2)
	dim errMsg
	dim i,j

	'Determine which services to get rates for
	if UPSallRates = "Y" then
		requestOption = "shop"
	else
		requestOption = "rate"
	end if
	errMsg = ""

	'Create XML request
	strXMLSend = "" & _
	"<?xml version=""1.0""?>" & _
	"<AccessRequest xml:lang=""en-US"">" & _
		"<AccessLicenseNumber>" & UPSACCESSID & "</AccessLicenseNumber>" & _
		"<UserId>" & UPSUSERID & "</UserId>" & _
		"<Password>" & UPSPASSWORD & "</Password>" & _
	"</AccessRequest>" & _
	"<?xml version=""1.0""?>" & _
	"<RatingServiceSelectionRequest xml:lang=""en-US"">" & _
		"<Request>" & _
			"<TransactionReference>" & _
				"<CustomerContext>Rating and Service</CustomerContext>" & _
				"<XpciVersion>1.0001</XpciVersion>" & _
			"</TransactionReference>" & _
			"<RequestAction>Rate</RequestAction>" & _
			"<RequestOption>" & REQUESTOPTION & "</RequestOption>" & _
		"</Request>" & _
		"<PickupType>" & _
			"<Code>" & UPSPICKUPTYPE & "</Code>" & _
		"</PickupType>" & _
		"<Shipment>" & _
			"<Shipper>" & _
				"<Address>" & _
					"<PostalCode>" & UPSFROMZIP & "</PostalCode>" & _
					"<CountryCode>" & UPSFROMCNTRY & "</CountryCode>" & _
				"</Address>" & _
			"</Shipper>" & _
			"<ShipTo>" & _
				"<Address>" & _
					"<PostalCode>" & UPSTOZIP & "</PostalCode>" & _
					"<CountryCode>" & UPSTOCOUNTRY & "</CountryCode>" & _
				"</Address>" & _
			"</ShipTo>" & _
			"<Service>" & _
				"<Code>" & UPSSHIPCODE & "</Code>" & _
			"</Service>" & _
			"<Package>" & _
				"<PackagingType>" & _
					"<Code>" & UPSPACKTYPE & "</Code>" & _
				"</PackagingType>" & _
				"<PackageWeight>" & _
					"<UnitOfMeasurement>" & _
						"<Code>" & UPSWEIGHTUNIT & "</Code>" & _
					"</UnitOfMeasurement>" & _
					"<Weight>" & UPSWEIGHT & "</Weight>" & _
				"</PackageWeight>" & _
			"</Package>" & _
			"<ShipmentServiceOptions/>" & _
		"</Shipment>" & _
	"</RatingServiceSelectionRequest>"

	'Send request
	on error resume next
	set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
	if err.number <> 0 then
		err.Clear
		set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP.4.0")
		if err.number <> 0 then
			UPSrates = "UPS : " & err.Description
			exit function
		end if
	end if
	on error goto 0
	xmlhttp.Open "POST","https://www.ups.com/ups.app/xml/Rate",false
	xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	xmlhttp.send strXMLSend
	if xmlhttp.status <> 200 then
		UPSrates = "UPS : " & xmlhttp.status & " - " & xmlhttp.statusText & "."
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
		errMsg = "UPS : Response from UPS could not be parsed."
	else
		set nodesError = xmlDom.documentElement.selectSingleNode("Response/ResponseStatusCode")
		if nodesError.Text <> "1" then
			errMsg = "UPS : " & xmlDom.documentElement.selectSingleNode("Response/Error/ErrorDescription").text
		else
			j = 0
			set nodesShipment = xmlDom.documentElement.selectNodes("RatedShipment")
			for each i in nodesShipment
				'Extract XML elements and data
				set nodesService  = i.selectSingleNode("Service/Code")
				set nodesRate     = i.selectSingleNode("TotalCharges/MonetaryValue")
				set nodesCurrency = i.selectSingleNode("TotalCharges/CurrencyCode")
				shipRateArr(j,0)  = nodesRate.text
				shipRateArr(j,1)  = UPSserviceDesc(nodesService.text) & " (" & nodesCurrency.text & ")"
				j = j + 1
			next
			'If no rates returned, then give error.
			if len(shipRateArr(0,0)) = 0 then
				errMsg = "UPS : No rates were returned."
			end if
		end if
	end if
	set xmlDom = nothing

	if len(errMsg) > 0 then
		UPSrates = errMsg
	else
		UPSrates = shipRateArr
	end if

end function

'UPS Service Descriptions
function UPSserviceDesc(serviceCode)
	select case serviceCode
		case "01"
			UPSserviceDesc = "UPS Next Day Air"
		case "02"
			UPSserviceDesc = "UPS 2nd Day Air"
		case "03"
			UPSserviceDesc = "UPS Ground"
		case "07"
			UPSserviceDesc = "UPS Worldwide Express"
		case "08"
			UPSserviceDesc = "UPS Worldwide Expedited"
		case "11"
			UPSserviceDesc = "UPS Standard"
		case "12"
			UPSserviceDesc = "UPS 3-Day Select"
		case "13"
			UPSserviceDesc = "UPS Next Day Air Saver"
		case "14"
			UPSserviceDesc = "UPS Next Day Air Early AM"
		case "54"
			UPSserviceDesc = "UPS Worldwide Express Plus"
		case "59"
			UPSserviceDesc = "UPS 2nd Day Air AM"
		case "65"
			UPSserviceDesc = "UPS Express Saver"
		case else
			UPSserviceDesc = ""
	end select
end function
%>
