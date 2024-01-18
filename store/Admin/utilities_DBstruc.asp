<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Check Physical structure of Database
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
'Declare variables
dim mySQL, cn, rs
dim i, result, field
dim fixSQL

'DataTypes 
Const adInteger		= "|3|"
Const adDouble		= "|5|"
Const adDBTimeStamp = "|7|135|"
Const adVarChar     = "|202|200|"
Const adLongVarChar = "|203|201|"

'Error Counter
dim errorCount
errorCount = 0

'Cater for SQL & Access
dim genAUTOINCREMENT
dim genMEMO
dim genTEXT
dim genTIMESTAMP
dim genDOUBLE

'*************************************************************************

if dbType = 1 then 'SQL Server
	genAUTOINCREMENT	= "INTEGER IDENTITY(1,1) PRIMARY KEY"
	genMEMO				= "TEXT"
	genTEXT				= "VARCHAR"
	genTIMESTAMP		= "DATETIME"
	genDOUBLE			= "FLOAT"
else
	'Microsoft.Jet.OLEDB.3.51 tweak for AutoIncrement values
	if instr(lCase(connString),lCase("Jet.OLEDB.3.51")) > 0 then
		genAUTOINCREMENT= "AUTOINCREMENT"
	else
		genAUTOINCREMENT= "AUTOINCREMENT(1,1) PRIMARY KEY"
	end if
	genMEMO				= "MEMO"
	genTEXT				= "TEXT"
	genTIMESTAMP		= "TIMESTAMP"
	genDOUBLE			= "DOUBLE"
end if

%>
<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Database Verification</font></b>
	<br><br>
</P>

<!-- Check Tables -->
<table border=0 cellspacing=0 cellpadding=10 width="100%" class="textBlock">
<tr><td>
<%

'Open DB
call openDb()

'*************************************************************************
'START : Table Validations
'(*,0) = Field Name
'(*,1) = Field Data Type
'(*,2) = Field Length  (-1=Not specified, -2=Auto Increment)
'(*,3) = Field Default (-1=Not specified)
'*************************************************************************

'Check storeAdmin

dim storeAdmin(100,3)
i = 0
addField storeAdmin,"idConfig",					adInteger,		-2,	-1
addField storeAdmin,"adminType",				adVarChar,		1,	-1
addField storeAdmin,"configVar",				adVarChar,		50,	-1
addField storeAdmin,"configVal",				adVarChar,		255,-1
addField storeAdmin,"configValLong",			adLongVarChar,	-1,	-1

checkTable storeAdmin, "storeAdmin"

'Check reviews

dim reviews(100,3)
i = 0
addField reviews,"idReview",					adInteger,		-2,	-1
addField reviews,"idProduct",					adInteger,		-1,	-1
addField reviews,"revDate",						adDBTimeStamp,	-1,	-1
addField reviews,"revDateInt",					adVarChar,		25,	-1
addField reviews,"revAuditInfo",				adVarChar,		255,-1
addField reviews,"revStatus",					adVarChar,		1,	-1
addField reviews,"revRating",					adInteger,		-1,	-1
addField reviews,"revName",						adVarChar,		255,-1
addField reviews,"revLocation",					adVarChar,		255,-1
addField reviews,"revEmail",					adVarChar,		100,-1
addField reviews,"revSubj",						adVarChar,		100,-1
addField reviews,"revDetail",					adLongVarChar,	-1, -1

checkTable reviews, "reviews"

'Check newsletters

dim newsletters(100,3)
i = 0
addField newsletters,"idNews",					adInteger,		-2,	-1
addField newsletters,"newsDate",				adDBTimeStamp,	-1,	-1
addField newsletters,"newsDateInt",				adVarChar,		25,	-1
addField newsletters,"newsBookmark",			adVarChar,		255,-1
addField newsletters,"newsSubj",				adVarChar,		255,-1
addField newsletters,"newsBody",				adLongVarChar,	-1, -1

checkTable newsletters, "newsletters"

'Check customer

dim customer(100,3)
i = 0
addField customer,"idCust",						adInteger,		-2,	-1
addField customer,"status",						adVarChar,		1,	-1
addField customer,"dateCreated",				adDBTimeStamp,	-1,	-1
addField customer,"dateCreatedInt",				adVarChar,		25,	-1
addField customer,"name",						adVarChar,		100,-1
addField customer,"lastName",					adVarChar,		100,-1
addField customer,"customerCompany",			adVarChar,		100,-1
addField customer,"phone",						adVarChar,		30,	-1
addField customer,"email",						adVarChar,		100,-1
addField customer,"password",					adVarChar,		80,	-1
addField customer,"address",					adVarChar,		100,-1
addField customer,"city",						adVarChar,		100,-1
addField customer,"locState",					adVarChar,		2,	-1
addField customer,"locState2",					adVarChar,		100,-1
addField customer,"locCountry",					adVarChar,		2,	-1
addField customer,"zip",						adVarChar,		20,	-1
addField customer,"paymentType",				adVarChar,		50,	-1
addField customer,"shippingName",				adVarChar,		100,-1
addField customer,"shippingLastName",			adVarChar,		100,-1
addField customer,"shippingPhone",				adVarChar,		30,	-1
addField customer,"shippingAddress",			adVarChar,		100,-1
addField customer,"shippingCity",				adVarChar,		100,-1
addField customer,"shippingLocState",			adVarChar,		2,	-1
addField customer,"shippingLocState2",			adVarChar,		100,-1
addField customer,"shippingLocCountry",			adVarChar,		2,	-1
addField customer,"shippingZip",				adVarChar,		20,	-1
addField customer,"futureMail",					adVarChar,		1,	-1
addField customer,"generalComments",			adLongVarChar,	-1,	-1
addField customer,"taxExempt",					adVarChar,		1,	"N"

checkTable customer, "customer"

'Check Products

dim products(100,3)
i = 0
addField products,"idProduct",					adInteger,		-2,	-1
addField products,"description",				adVarChar,		250,-1
addField products,"descriptionLong",			adVarChar,		250,-1
addField products,"details",					adLongVarChar,	-1,	-1
addField products,"relatedKeys",				adVarChar,		250,-1
addField products,"price",						adDouble,		-1,	-1
addField products,"listPrice",					adDouble,		-1,	-1
addField products,"imageUrl",					adVarChar,		50,	-1
addField products,"smallImageUrl",				adVarChar,		50,	-1
addField products,"sku",						adVarChar,		16,	-1
addField products,"stock",						adInteger,		-1,	-1
addField products,"weight",						adDouble,		-1,	-1
addField products,"active",						adInteger,		-1,	-1
addField products,"hotDeal",					adInteger,		-1,	-1
addField products,"homepage",					adInteger,		-1,	-1
addField products,"fileName",					adVarChar,		250,-1
addField products,"noShipCharge",				adVarChar,		1,	-1
addField products,"taxExempt",					adVarChar,		1,	"N"
addField products,"reviewAllow",				adVarChar,		1,	"N"
addField products,"reviewAutoActive",			adVarChar,		1,	"N"

checkTable products, "products"

'Check productGroups

dim productGroups(100,3)
i = 0
addField productGroups,"idProdGroup",			adInteger,		-2,	-1
addField productGroups,"prodGroupP",			adInteger,		-1,	-1
addField productGroups,"prodGroupC",			adInteger,		-1,	-1

checkTable productGroups, "productGroups"

'Check CartHead

dim carthead(100,3)
i = 0
addField carthead,"idOrder",					adInteger,		-2,	-1
addField carthead,"idCust",						adInteger,		-1,	-1
addField carthead,"orderDate",					adDBTimeStamp,	-1,	-1
addField carthead,"orderDateInt",				adVarChar,		25,	-1
addField carthead,"randomKey",					adVarChar,		50,	-1
addField carthead,"subTotal",					adDouble,		-1,	-1
addField carthead,"taxTotal",					adDouble,		-1,	-1
addField carthead,"shipmentTotal",				adDouble,		-1,	-1
addField carthead,"handlingFeeTotal",			adDouble,		-1,	0
addField carthead,"Total",						adDouble,		-1,	-1
addField carthead,"shipmentMethod",				adVarChar,		100,-1
addField carthead,"name",						adVarChar,		100,-1
addField carthead,"lastName",					adVarChar,		100,-1
addField carthead,"customerCompany",			adVarChar,		100,-1
addField carthead,"phone",						adVarChar,		30,	-1
addField carthead,"email",						adVarChar,		100,-1
addField carthead,"address",					adVarChar,		100,-1
addField carthead,"city",						adVarChar,		100,-1
addField carthead,"locState",					adVarChar,		100,-1
addField carthead,"locCountry",					adVarChar,		100,-1
addField carthead,"zip",						adVarChar,		20,	-1
addField carthead,"shippingName",				adVarChar,		100,-1
addField carthead,"shippingLastName",			adVarChar,		100,-1
addField carthead,"shippingPhone",				adVarChar,		30,	-1
addField carthead,"shippingAddress",			adVarChar,		100,-1
addField carthead,"shippingCity",				adVarChar,		100,-1
addField carthead,"shippingLocState",			adVarChar,		100,-1
addField carthead,"shippingLocCountry",			adVarChar,		100,-1
addField carthead,"shippingZip",				adVarChar,		20,	-1
addField carthead,"paymentType",				adVarChar,		50,	-1
addField carthead,"cardType",					adVarChar,		50,	-1
addField carthead,"cardNumber",					adVarChar,		50,	-1
addField carthead,"cardExpMonth",				adVarChar,		2,	-1
addField carthead,"cardExpYear",				adVarChar,		4,	-1
addField carthead,"cardVerify",					adVarChar,		4,	-1
addField carthead,"cardName",					adVarChar,		100,-1
addField carthead,"generalComments",			adVarChar,		255,-1
addField carthead,"orderStatus",				adVarChar,		1,	-1
addField carthead,"auditInfo",					adVarChar,		255,-1
addField carthead,"storeComments",				adLongVarChar,	-1,	-1
addField carthead,"storeCommentsPriv",			adLongVarChar,	-1,	-1
addField carthead,"adjustAmount",				adDouble,		-1,	-1
addField carthead,"adjustReason",				adVarChar,		255,-1
addField carthead,"taxExempt",					adVarChar,		1,	"N"
addField carthead,"discCode",					adVarChar,		20,	-1
addField carthead,"discPerc",					adDouble,		-1,	-1
addField carthead,"discTotal",					adDouble,		-1,	-1

checkTable carthead, "carthead"

'Check CartRows

dim cartrows(100,3)
i = 0
addField cartrows,"idCartRow",					adInteger,		-2,	-1
addField cartrows,"idOrder",					adInteger,		-1,	-1
addField cartrows,"idProduct",					adInteger,		-1,	-1
addField cartrows,"sku",						adVarChar,		16,	-1
addField cartrows,"quantity",					adInteger,		-1,	-1
addField cartrows,"unitPrice",					adDouble,		-1,	-1
addField cartrows,"unitWeight",					adDouble,		-1,	-1
addField cartrows,"description",				adVarChar,		250,-1
addField cartrows,"downloadCount",				adInteger,		-1,	-1
addField cartrows,"downloadDate",				adVarChar,		25,	-1
addField cartrows,"taxExempt",					adVarChar,		1,	"N"
addField cartrows,"idDiscProd",					adInteger,		-1,	-1
addField cartrows,"discAmt",					adDouble,		-1,	-1

checkTable cartrows, "cartrows"

'Check CartRowsOptions

dim CartRowsOptions(100,3)
i = 0
addField CartRowsOptions,"idCartRowOption",		adInteger,		-2,	-1
addField CartRowsOptions,"idOrder",				adInteger,		-1,	-1
addField CartRowsOptions,"idCartRow",			adInteger,		-1,	-1
addField CartRowsOptions,"idOption",			adInteger,		-1,	-1
addField CartRowsOptions,"optionPrice",			adDouble,		-1,	-1
addField CartRowsOptions,"optionWeight",		adDouble,		-1,	-1
addField CartRowsOptions,"optionDescrip",		adVarChar,		255,-1
addField CartRowsOptions,"taxExempt",			adVarChar,		1,	"N"

checkTable CartRowsOptions, "CartRowsOptions"

'Check Categories

dim Categories(100,3)
i = 0
addField Categories,"idCategory",				adInteger,		-2,	-1
addField Categories,"categoryDesc",				adVarChar,		50,	-1
addField Categories,"idParentCategory",			adInteger,		-1,	-1
addField Categories,"categoryFeatured",			adVarChar,		 1,	"N"
addField Categories,"categoryHTML",				adVarChar,		255,-1

checkTable Categories, "Categories"

'Check Categories_Products

dim Categories_Products(100,3)
i = 0
addField Categories_Products,"idCatProd",		adInteger,		-2,	-1
addField Categories_Products,"idProduct",		adInteger,		-1,	-1
addField Categories_Products,"idCategory",		adInteger,		-1,	-1

checkTable Categories_Products, "Categories_Products"

'Check Locations

dim Locations(100,3)
i = 0
addField Locations,"idLocation",				adInteger,		-2,	-1
addField Locations,"locName",					adVarChar,		100,-1
addField Locations,"locCountry",				adVarChar,		2,	-1
addField Locations,"locState",					adVarChar,		2,	-1
addField Locations,"locTax",					adDouble,		-1,	-1
addField Locations,"locShipZone",				adInteger,		-1,	-1
addField Locations,"locStatus",					adVarChar,		1,	"A"

checkTable Locations, "Locations"

'Check OptionsProdEx

dim OptionsProdEx(100,3)
i = 0
addField OptionsProdEx,"idOptionsProdEx",		adInteger,		-2,	-1
addField OptionsProdEx,"idOption",				adInteger,		-1,	-1
addField OptionsProdEx,"idProduct",				adInteger,		-1,	-1

checkTable OptionsProdEx, "OptionsProdEx"

'Check Options

dim Options(100,3)
i = 0
addField Options,"idOption",					adInteger,		-2,	-1
addField Options,"optionDescrip",				adVarChar,		50,	-1
addField Options,"priceToAdd",					adDouble,		-1,	-1
addField Options,"weightToAdd",					adDouble,		-1,	-1
addField Options,"taxExempt",					adVarChar,		 1,	"N"
addField Options,"percToAdd",					adDouble,		-1,	0

checkTable Options, "Options"

'Check optionsXref

dim optionsXref(100,3)
i = 0
addField optionsXref,"idOptOptGroup",			adInteger,		-2,	-1
addField optionsXref,"idOptionGroup",			adInteger,		-1,	-1
addField optionsXref,"idOption",				adInteger,		-1,	-1

checkTable optionsXref, "optionsXref"

'Check OptionsGroups

dim OptionsGroups(100,3)
i = 0
addField OptionsGroups,"idOptionGroup",			adInteger,		-2,	-1
addField OptionsGroups,"optionGroupDesc",		adVarChar,		50,	-1
addField OptionsGroups,"optionReq",				adVarChar,		1,	-1
addField OptionsGroups,"optionType",			adVarChar,		1, "S"

checkTable OptionsGroups, "OptionsGroups"

'Check optionsGroupsXref

dim optionsGroupsXref(100,3)
i = 0
addField optionsGroupsXref,"idOptGrpProd",		adInteger,		-2,	-1
addField optionsGroupsXref,"idProduct",			adInteger,		-1,	-1
addField optionsGroupsXref,"idOptionGroup",		adInteger,		-1,	-1

checkTable optionsGroupsXref, "optionsGroupsXref"

'Check ShipMethod

dim ShipMethod(100,3)
i = 0
addField ShipMethod,"idShipMethod",				adInteger,		-2,	-1
addField ShipMethod,"shipDesc",					adVarChar,		100,-1
addField ShipMethod,"status",					adVarChar,		1,	-1

checkTable ShipMethod, "ShipMethod"

'Check ShipRates

dim ShipRates(100,3)
i = 0
addField ShipRates,"idShip",					adInteger,		-2,	-1
addField ShipRates,"locShipZone",				adInteger,		-1,	-1
addField ShipRates,"idShipMethod",				adInteger,		-1,	-1
addField ShipRates,"unitType",					adVarChar,		1,	-1
addField ShipRates,"unitsFrom",					adDouble,		-1,	-1
addField ShipRates,"unitsTo",					adDouble,		-1,	-1
addField ShipRates,"addAmt",					adDouble,		-1,	-1
addField ShipRates,"addPerc",					adDouble,		-1,	-1

checkTable ShipRates, "ShipRates"

'Check DiscOrder

dim DiscOrder(100,3)
i = 0
addField DiscOrder,"idDiscOrder",				adInteger,		-2,	-1
addField DiscOrder,"discCode",					adVarChar,		20,	-1
addField DiscOrder,"discPerc",					adDouble,		-1,	-1
addField DiscOrder,"discAmt",					adDouble,		-1,	-1
addField DiscOrder,"discFromAmt",				adDouble,		-1,	-1
addField DiscOrder,"discToAmt",					adDouble,		-1,	-1
addField DiscOrder,"discStatus",				adVarChar,		 1,	-1
addField DiscOrder,"discOnceOnly",				adVarChar,		 1,	-1
addField DiscOrder,"discValidFrom",				adVarChar,		25,	-1
addField DiscOrder,"discValidTo",				adVarChar,		25,	-1

checkTable DiscOrder, "DiscOrder"

'Check DiscProd

dim DiscProd(100,3)
i = 0
addField DiscProd,"idDiscProd",					adInteger,		-2,	-1
addField DiscProd,"discAmt",					adDouble,		-1,	-1
addField DiscProd,"discFromQty",				adDouble,		-1,	-1
addField DiscProd,"discToQty",					adDouble,		-1,	-1
addField DiscProd,"idProduct",					adInteger,		-1,	-1
addField DiscProd,"discPerc",					adDouble,		-1,	-1

checkTable DiscProd, "DiscProd"

'*************************************************************************
'END : Table Validations
'*************************************************************************
%>
</td></tr>
</table>

<br>

<!-- Status -->
<table border=0 cellspacing=0 cellpadding=10 width="100%" class="textBlock">
<tr><td>
<%
if errorCount > 0 then
%>
	<br><b><font size=2 color=red><%=errorCount%> error(s) were found!</font></b><br><br>
<%
	if len(trim(fixSQL)) > 0 then
%>
		<font size=2>
			<b>NOTE :</b> To fix your database, you will need to run 
			the repair script below by clicking on the "Run&nbsp;Fix&nbsp;Now" 
			button. The database will automatically be re-tested after 
			all the fixes have been applied. If errors persist, then 
			fixes will have to be applied manually.
		</font>
		
		<br>
		
		<form method="post" action="utilities_DBstrucExec.asp" name="fixDB">
			<input type=hidden name=fixSQL value="<%=fixSQL%>">
			<input type=submit name=submit value="Run Fix Now">
		</form>

		<table border=0 cellspacing=0 cellpadding=8 class="blockInBlock">
			<tr><td><pre><%=replace(fixSQL,"*|*",vbCRlf&vbCRlf)%></pre></td></tr>
		</table>
		
		<br><br>
<%
	else
%>
		<font size=2>
			<b>NOTE :</b> Please correct the indicated errors manually. 
			If you are unsure what the field definitions should be (ie. 
			length, type, etc.), refer to the example MS Access database 
			that came with this package. ALWAYS re-run this test after 
			you have made your changes to make sure that the changes 
			were applied correctly.
		</font>
		<br><br>
<%
	end if
else
%>
	<br><b><font size=2 color=green>Congratulations! No errors were found.</font></b><br><br>
<%
end if

'Close Database
call closedb()

%>
</td></tr>
</table>

<!--#include file="_INCfooter_.asp"-->
<%
'*************************************************************************
'Add another field definition to a Table's field definition array
'*************************************************************************
sub addField(tableArray,fieldName,fieldType,fieldLength,fieldDefault)
	tableArray(i,0) = fieldName
	tableArray(i,1) = fieldType
	tableArray(i,2) = fieldLength
	tableArray(i,3) = fieldDefault
	i = i + 1
end sub
'*************************************************************************
'Check a Table
'*************************************************************************
sub checkTable(tableArray,tableName)
	Response.Write "<font size=2>Checking Table --> </font><b><font size=2 color=#800000>" & tableName & "</font></b><br><br>"
	on error resume next
	set rs = server.CreateObject("adodb.recordset")
	rs.Open "SELECT TOP 1 *,'dummy' AS dummyField FROM " & tableName, cn
	if Err.number <> 0 then
		on error goto 0
		errorCount = errorCount + 1
		Response.Write "<b><font color=red>Table Not Found!</font></b><br>"
		result = createTableSQL(tableArray,tableName)
	else
		on error goto 0
		'Check fields in tableArray against DB table fields
		for i = 0 to UBound(tableArray)
			if len(trim(tableArray(i,0))) > 0 then
				select case checkField(rs,tableArray(i,0),tableArray(i,1),tableArray(i,2))
					case 0
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<img src=x_Y.gif border=0 valign=absMiddle> <b>" & tableArray(i,0) & "</b>"
					case 1
						errorCount = errorCount + 1
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<img src=x_N.gif border=0 valign=absMiddle> <b>" & tableArray(i,0) & "</b>" & " - "
						Response.Write "<font color=red>Field not found.</font>"
						result = createFieldSQL(tableArray(i,0),tableArray(i,1),tableArray(i,2),tableArray(i,3),tableName)
					case 2
						errorCount = errorCount + 1
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<img src=x_N.gif border=0 valign=absMiddle> <b>" & tableArray(i,0) & "</b>" & " - "
						Response.Write "<font color=red>Field Type Invalid.</font>"
						result = modifyFieldSQL(tableArray(i,0),tableArray(i,1),tableArray(i,2),tableArray(i,3),tableName)
					case 3
						errorCount = errorCount + 1
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<img src=x_N.gif border=0 valign=absMiddle> <b>" & tableArray(i,0) & "</b>" & " - "
						Response.Write "<font color=red>Field Length Invalid.</font>"
						result = modifyFieldSQL(tableArray(i,0),tableArray(i,1),tableArray(i,2),tableArray(i,3),tableName)
					case 4
						errorCount = errorCount + 1
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<img src=x_N.gif border=0 valign=absMiddle> <b>" & tableArray(i,0) & "</b>" & " - "
						Response.Write "<font color=red>Field should be AutoNumber/Increment.</font>"
					case else
						response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Unknown Field Match Type returned. Script could not continue.")
				end select
				Response.Write "<br>"
			end if
		next
		'Check DB Table for fields that are not required
		for each field in rs.fields
			if lCase(field.Name) <> "dummyfield" then 'Field used to force the return of a recordset even if Table is empty
				for i = 0 to UBound(tableArray)
					if lCase(field.name) = lCase(tableArray(i,0)) then
						exit for
					end if
				next
				if i = Ubound(tableArray) + 1 then
					errorCount = errorCount + 1
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<img src=x_N.gif border=0 valign=absMiddle> <b>" & field.name & "</b>" & " - "
					Response.Write "<font color=red>Field is not Required.</font>"
					Response.Write "<br>"
				end if
			end if
		next
		rs.Close
		set rs = nothing
	end if
	Response.Write "<br><br>"
end sub
'*************************************************************************
'Check a recordset to see if there is a field in there that matches where : 
'field in recordset = fieldName + fieldType + fieldLength.
'*************************************************************************
function checkField(rs,fieldName,fieldType,fieldLength)

	'Returns:
	'0 - If complete match is found
	'1 - If Field Name was not found
	'2 - If Field Type is invalid
	'3 - If Field Length is invalid
	'4 - If Field should be AutoNumber but is not
		
	dim field
	
	for each field in rs.Fields
		if lCase(field.name) <> lCase(fieldName) then
			checkField = 1
		else
			checkField = 0
			if instr(fieldType,"|" & Cstr(field.Type) & "|") = 0 then
				checkField = 2
			else
				if field.DefinedSize <> fieldLength then
					if fieldLength > 0 then
						checkField = 3
					else
						'If -2 we test for AutoNumber field
						if fieldLength = -2 then
							If not field.properties("ISAUTOINCREMENT") Then
								checkField = 4
							end if
						end if
					end if
				end if
			end if
			exit for
		end if
	next
	
end function
'*************************************************************************
'Create SQL Statement to CREATE a Table
'*************************************************************************
function createTableSQL(tableArray,tableName)

	dim i
	
	fixSQL = fixSQL & "CREATE TABLE " & tableName & " (" & vbCRLf
	for i = 0 to UBound(tableArray)
		if len(trim(tableArray(i,0))) > 0 then
			if tableArray(i,2) = -2 then
				fixSQL = fixSQL & tableArray(i,0) & " " & genAUTOINCREMENT
			else
				select case tableArray(i,1)
					case adInteger
						fixSQL = fixSQL & tableArray(i,0) & " INTEGER"
					case adDouble
						fixSQL = fixSQL & tableArray(i,0) & " " & genDOUBLE
					case adVarChar
						fixSQL = fixSQL & tableArray(i,0) & " " & genTEXT & "(" & tableArray(i,2) & ")"
					case adDBTimeStamp
						fixSQL = fixSQL & tableArray(i,0) & " " & genTIMESTAMP
					case adLongVarChar
						fixSQL = fixSQL & tableArray(i,0) & " " & genMEMO
					case else
				end select
			end if
			fixSQL = fixSQL & "," & vbCRLf
		end if
	next
	fixSQL = mid(fixSQL,1,len(fixSQL)-3) 'Get rid of last comma
	fixSQL = fixSQL & ")"
	fixSQL = fixSQL & "*|*" 'End of SQL Sentence
	
end function
'*************************************************************************
'Create SQL Statement to ADD a Column
'*************************************************************************
function createFieldSQL(fieldName,fieldType,fieldLength,fieldDefault,tableName)

	'Generate ADD COLUMN code
	fixSQL = fixSQL & "ALTER TABLE " & tableName & " ADD "
	select case fieldType
		case adInteger
			fixSQL = fixSQL & fieldName & " INTEGER"
		case adDouble
			fixSQL = fixSQL & fieldName & " " & genDOUBLE
		case adVarChar
			fixSQL = fixSQL & fieldName & " " & genTEXT & "(" & fieldLength & ")"
		case adDBTimeStamp
			fixSQL = fixSQL & fieldName & " " & genTIMESTAMP
		case adLongVarChar
			fixSQL = fixSQL & fieldName & " " & genMEMO
		case else
	end select
	
	'Append statement to populate new Column with a value (if required)
	if fieldDefault <> -1 then
		select case fieldType
			case adInteger, adDouble
				fixSQL = fixSQL & "*|*" & "UPDATE " & tableName & " SET " & fieldName & "=" & fieldDefault
			case adVarChar, adLongVarChar
				fixSQL = fixSQL & "*|*" & "UPDATE " & tableName & " SET " & fieldName & "='" & fieldDefault & "'"
			case else
		end select
	end if
	
	'Mark the end of the SQL sentence
	fixSQL = fixSQL & "*|*"
	
end function
'*************************************************************************
'Create SQL Statement to ALTER a Column
'*************************************************************************
function modifyFieldSQL(fieldName,fieldType,fieldLength,fieldDefault,tableName)

	'Generate ALTER COLUMN code
	fixSQL = fixSQL & "ALTER TABLE " & tableName & " ALTER COLUMN "
	select case fieldType
		case adInteger
			fixSQL = fixSQL & fieldName & " INTEGER"
		case adDouble
			fixSQL = fixSQL & fieldName & " " & genDOUBLE
		case adVarChar
			fixSQL = fixSQL & fieldName & " " & genTEXT & "(" & fieldLength & ")"
		case adDBTimeStamp
			fixSQL = fixSQL & fieldName & " " & genTIMESTAMP
		case adLongVarChar
			fixSQL = fixSQL & fieldName & " " & genMEMO
		case else
	end select
	
	'Mark the end of the SQL sentence
	fixSQL = fixSQL & "*|*"
	
end function
%>