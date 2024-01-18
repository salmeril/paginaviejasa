<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : These functions and subroutines are used by the scripts
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************

'*************************************************************************
'Calculate cart total
'INCLUDE - Product (Item) Discounts
'INCLUDE - Order Discount
'EXCLUDE - Taxes and Shipping
'*************************************************************************
function cartTotal(idOrder,idCartRow)

	'Declare Variables
	dim mySQL,rsTemp
	dim discPerc
	dim discTotal
	cartTotal = 0.00
	
	'Calculate totals
	if  len(idOrder)   > 0 and IsNumeric(idOrder) _
	and len(idCartRow) > 0 and IsNumeric(idCartRow) then
		mySQL = "SELECT discPerc " _
		      & "FROM   cartHead " _
		      & "WHERE  idOrder = " & validSQL(idOrder,"I")
		set rsTemp = openRSexecute(mySQL)
		if not rsTemp.EOF then
			discPerc  = CDbl(emptyString(rsTemp("discPerc"),"0"))
			cartTotal = cartTotalExDisc(idOrder,idCartRow)
			discTotal = Round(((cartTotal * discPerc) / 100),2)
			cartTotal = cartTotal - discTotal
		end if
		call closeRS(rsTemp)
	end if

	cartTotal = CDbl(cartTotal)
		  
end function

'*************************************************************************
'Calculate cart total
'INCLUDE - Product (Item) Discounts
'EXCLUDE - Order Discount
'EXCLUDE - Taxes and Shipping
'*************************************************************************
function cartTotalExDisc(idOrder,idCartRow)

	'Declare Variables
	dim mySQL,rsTemp
	dim quantity,unitPrice
	dim optionPrice,discAmt
	cartTotalExDisc = 0.00
	
	'Calculate totals
	if  len(idOrder)   > 0 and IsNumeric(idOrder) _
	and len(idCartRow) > 0 and IsNumeric(idCartRow) then
		mySQL = "SELECT quantity,unitPrice,discAmt," _
			  & "       (SELECT SUM(optionPrice) " _
			  & "       FROM    cartRowsOptions " _
			  & "       WHERE   cartRowsOptions.idCartRow = cartRows.idCartRow) " _
			  & "       AS      optionPrice " _
			  & "FROM   cartRows " _
			  & "WHERE  idOrder = " & validSQL(idOrder,"I") & " "
		if idCartRow > 0 then
			mySQL = mySQL & " AND idCartRow = " & validSQL(idCartRow,"I")
		end if
		set rsTemp = openRSexecute(mySQL)
		do while not rsTemp.eof
			quantity    = CDbl(emptyString(rsTemp("quantity"),"0"))
			unitPrice   = CDbl(emptyString(rsTemp("unitPrice"),"0"))
			discAmt     = CDbl(emptyString(rsTemp("discAmt"),"0"))
			optionPrice = CDbl(emptyString(rsTemp("optionPrice"),"0"))
			cartTotalExDisc = cartTotalExDisc + (quantity * (unitPrice + optionPrice - discAmt))
			rsTemp.movenext
		loop
		call closeRS(rsTemp)
	end if
	
	cartTotalExDisc = CDbl(cartTotalExDisc)
		  
end function

'*************************************************************************
'Calculate Cart Quantity
'*************************************************************************
function cartQty(idOrder)

	'Declare Variables
	dim mySQL, rsTemp
	cartQty = 0.00
	
	'Calculate Cart Quantity
	if len(idOrder) > 0 and IsNumeric(idOrder) then
		mySQL = "SELECT SUM(quantity) AS qTotal " _
			  & "FROM   cartRows " _
			  & "WHERE  idOrder = " & validSQL(idOrder,"I")
		set rsTemp = openRSexecute(mySQL)
		if not rsTemp.eof then
			if isNumeric(rsTemp("qTotal")) then
				cartQty = rsTemp("qTotal")
			end if
		end if
		call closeRS(rsTemp)
	end if
	
	cartQty = CDbl(cartQty)
	 
end function
'*************************************************************************
'Money Formatter - Use Store LCID
'*************************************************************************
function moneyS(aNumber)
	if isNumeric(aNumber) then
		dim tempNumber
		tempNumber   = CDbl(aNumber)
		session.LCID = pStoreLCID	'User configured format
		moneyS		 = formatNumber(round(tempNumber,2),2)
		session.LCID = 1033			'Default format
	else
		moneyS = aNumber
	end if
end function
'*************************************************************************
'Money Formatter - Use Default LCID
'*************************************************************************
function moneyD(aNumber)
	if isNumeric(aNumber) then
		moneyD = formatnumber(round(aNumber,2),2)
	else
		moneyD = aNumber
	end if
end function
'*************************************************************************
'Date formatter
'*************************************************************************
function formatTheDate(strDate)
	if isDate(strDate) then
		dim tempDate
		tempDate	  = CDate(strDate)
		session.LCID  = pStoreLCID	'User configured format
		formatTheDate = formatDateTime(tempDate,vbShortDate)
		session.LCID  = 1033		'Default format
	else
		formatTheDate = strDate
	end if
end function
'*************************************************************************
'Scan list of error FieldNames for possible match
'*************************************************************************
function checkFieldError(byVal FieldName, array1)
	dim i
	FieldName = Lcase(FieldName)
	for i = 0 to Ubound(array1)
		if LCase(array1(i)) = FieldName then
			checkFieldError = "<font color=red>*</font>"
			exit for
		end if
	next
end function
'*************************************************************************
'Substitute empty or null strings with something else
'*************************************************************************
function emptyString(tempStr,replaceWith)
	if len(trim(tempStr))=0 or isEmpty(tempStr) or isNull(tempStr) then
		emptyString = replaceWith
	else
		emptyString = trim(tempStr)
	end if
end function
'*************************************************************************
'Payment Type Message / Description
'*************************************************************************
function paymentMsg(paymentType,Amount,cardNumber)
	if Amount > 0 then
		select case lCase(paymentType)
		case "mailin"
			paymentMsg = payMsgMailIn
		case "callin"
			paymentMsg = payMsgCallIn
		case "faxin"
			paymentMsg = payMsgFaxIn
		case "cod"
			paymentMsg = payMsgCOD
		case "creditcard"
			if len(trim(cardNumber)) > 4 then
				paymentMsg = payMsgCreditCard & " (" & replace(space(len(cardNumber)-4)," ","x") & right(cardNumber,4) & ")"
			else
				paymentMsg = payMsgCreditCard
			end if
		case "paypal"
			paymentMsg = payMsgPayPal
		case "2checkout"
			paymentMsg = payMsgTwoCheckOut
		case "authorizenet"
			paymentMsg = payMsgAuthNet
		case "custom"
			paymentMsg = payMsgCustom
		case else
			paymentMsg = payMsgOther
		end select
	else
		paymentMsg = payMsgNotReq
	end if
end function
'******************************************************************
'Get/Set idOrder from session/form/querystring
'******************************************************************
function sessionCart()

	'Declare Variables
	dim mySQL, rsTemp, idOrder
	idOrder = trim(session(storeID & "idOrder"))
	
	'Check idOrder exists and Order is still Open
	if isEmpty(idOrder) or not IsNumeric(idOrder) then
		session(storeID & "idOrder") = null
		sessionCart					 = null
	else	
		mySQL="SELECT idOrder " _
		    & "FROM   cartHead " _
		    & "WHERE  idOrder = " & validSQL(idOrder,"I") & " " _
		    & "AND   (orderStatus = 'U' OR orderStatus = 'S') "
		set rsTemp = openRSexecute(mySQL)
		if not rstemp.eof then
			session(storeID & "idOrder") = idOrder
			sessionCart					 = idOrder
		else
			session(storeID & "idOrder") = null
			sessionCart					 = null
		end if
		call closeRS(rsTemp)
	end if
	
end function
'******************************************************************
'Get/Set idCust from session/form/querystring
'******************************************************************
function sessionCust()

	'Declare Variables
	dim mySQL, rsTemp, idCust
	idCust = trim(session(storeID & "idCust"))
	
	'Check if idCust exists on DB and is still Active
	if isEmpty(idCust) or not IsNumeric(idCust) then
		session(storeID & "idCust")	= null
		sessionCust					= null
	else
		mySQL="SELECT idCust FROM customer " _
		    & "WHERE  idCust = " & validSQL(idCust,"I") & " " _
		    & "AND    status = 'A'"
		set rsTemp = openRSexecute(mySQL)
		if not rstemp.eof then
			session(storeID & "idCust")	= idCust
			sessionCust					= idCust
		else
			session(storeID & "idCust")	= null
			sessionCust					= null
		end if
		call closeRS(rsTemp)
	end if
		
end function
'******************************************************************
'Format values entered into HTML form fields to prevent cross-site 
'scripting and other malicious HTML.
'******************************************************************
function validHTML(aString)

	'Declare Variables
	dim tempString
	tempString = trim(aString)
	
	'Check for empty values
	if isNull(tempString) or isEmpty(tempString) or len(tempString) = 0 then
		validHTML = ""
		exit function
	end if

	'Clean up HTML
	tempString = replace(tempString,"<", " ")
	tempString = replace(tempString,">", " ")
	tempString = replace(tempString,"""","'")
	validHTML  = trim(tempString)
	
end function
'******************************************************************
'Format values inserted into SQL statements before executing the 
'SQL statement. This is to prevent SQL injection attacks, and to 
'ensure that certain characters are interpreted correctly.
'******************************************************************
function validSQL(aString,aType)

	'Declare Variables
	dim tempString
	tempString = trim(aString)
	
	'Check for empty values
	if isNull(tempString) or isEmpty(tempString) or len(tempString) = 0 then
		validSQL = ""
		exit function
	end if
	
	'Clean up SQL
	if lCase(tempString) = "null" then				'Nulls
		validSQL = tempString
	else
		select case trim(UCase(aType))
		case "I"									'Integer
			validSQL = CLng(tempString)
		case "D"									'Double
			validSQL = CDbl(tempString)
		case else									'Alphanumeric
			tempString = replace(tempString,"--"," ")
			tempString = replace(tempString,"=="," ")
			tempString = replace(tempString,";", " ")
			tempString = replace(tempString,"'","''")
			validSQL   = tempString
		end select
	end if
	
end function
'******************************************************************
'Check a string for invalid characters
'******************************************************************
function invalidChar(aString,alphaNum,addChars)

	dim i, checkChar

	invalidChar = true 'Assume invalid chars unless proven otherwise

	select case alphaNum
		case 1		'Alphanumeric [a-z, 0-9] is valid
			addChars = lCase("abcdefghijklmnopqrstuvwxyz0123456789" & addChars)
		case 2		'Numeric [0-9] is valid
			addChars = lCase("0123456789" & addChars)
		case 3		'Alpha [a-z] is valid
			addChars = lCase("abcdefghijklmnopqrstuvwxyz" & addChars)
		case else	'Only characters in addChar is valid
	end select
	
	for i = 1 to len(aString)
		checkChar = lCase(mid(aString,i,1))
		if inStr(addChars,checkChar) = 0 then
			invalidChar = true
			exit function
		end if
	next

	invalidChar = false
		
end function
'******************************************************************
'Convert Date to Integer
'******************************************************************
function dateInt(strDate)

	dim qYear, qMonth, qDay, qHour, qMin, qSec
	
	qYear   = year(strDate)
	qMonth  = left("00",2-len(datePart("m",strDate))) & datePart("m",strDate)
	qDay    = left("00",2-len(datePart("d",strDate))) & datePart("d",strDate)
	qHour   = left("00",2-len(datePart("h",strDate))) & datePart("h",strDate)
	qMin    = left("00",2-len(datePart("n",strDate))) & datePart("n",strDate)
	qSec    = left("00",2-len(datePart("s",strDate))) & datePart("s",strDate)
	
	dateInt = qYear & qMonth & qDay & qHour & qMin & qSec

end function
'******************************************************************
'Order Status Descriptions
'******************************************************************
function orderStatusDesc(orderStatus)
	select case orderStatus
	case "U"
		orderStatusDesc = langGenStatUnfinal
	case "S"
		orderStatusDesc = langGenStatSaved
	case "0"
		orderStatusDesc = langGenStatPending
	case "1"
		orderStatusDesc = langGenStatPaid
	case "2"
		orderStatusDesc = langGenStatShipped
	case "7"
		orderStatusDesc = langGenStatComplete
	case "9"
		orderStatusDesc = langGenStatCancel
	case else
		orderStatusDesc = langGenStatUnknown
	end select	
end function
'*************************************************************************
'Get State Description
'*************************************************************************
function getStateDesc(locCountry,locState,locState2)

	'Declare Variables
	dim mySQL, rsTemp
	locCountry = trim(locCountry)
	locState   = trim(locState)
	locState2  = trim(locState2)
	
	'If the alternate state is entered, return it.
	if len(locState2) > 0 then
		getStateDesc = locState2
	else
		'Get State description on database.
		if len(locCountry) = 0 or len(locState) = 0 then
			getStateDesc = locState
		else
			'Get State Name
			mySQL = "SELECT locName " _
			      & "FROM   locations " _
			      & "WHERE  locCountry = '" & validSQL(locCountry,"A") & "' " _
			      & "AND    locState = '"   & validSQL(locState,"A")   & "'"
			set rsTemp = openRSexecute(mySQL)
			if rsTemp.eof then
				getStateDesc = locState
			else
				getStateDesc = rsTemp("locName")
			end if
			call closeRS(rsTemp)
		end if
	end if
			
end function
'*************************************************************************
'Get Country Description
'*************************************************************************
function getCountryDesc(locCountry)

	'Declare Variables
	dim mySQL, rsTemp
	locCountry = trim(locCountry)
	
	'Check Country code
	if len(locCountry) = 0 then
		getCountryDesc = locCountry
	else
		'Get Country Name
		mySQL = "SELECT locName " _
		      & "FROM   locations " _
		      & "WHERE  locCountry = '" & validSQL(locCountry,"A") & "' " _
		      & "AND   (locState = '' OR locState IS NULL)"
		set rsTemp = openRSexecute(mySQL)
		if rsTemp.eof then
			getCountryDesc = locCountry
		else
			getCountryDesc = rsTemp("locName")
		end if
		call closeRS(rsTemp)
	end if
	
end function
'*************************************************************************
'Check if an Item is a Downloadable Item. If it is, return the filename
'of the downloadable file.
'*************************************************************************
function downloadFile(qIdOrder,idCartRow)

	'Declare Variables
	dim mySQL, rsTemp

	'Get Filename
	mySQL="SELECT products.fileName " _
		& "FROM   cartRows, products " _
		& "WHERE  idOrder = "   & validSQL(qIdOrder,"I")  & " " _
		& "AND    idCartRow = " & validSQL(idCartRow,"I") & " " _
		& "AND    products.idProduct = cartRows.idProduct " _
		& "AND    NOT (products.fileName IS NULL " _
		& "OR     products.fileName = '') "
	set rsTemp = openRSexecute(mySQL)
	if rsTemp.eof then
		downloadFile = ""
	else
		downloadFile = trim(rsTemp("fileName"))
	end if
	call closeRS(rsTemp)
	
end function
'*********************************************************************
'Check if str1 and str2 matches and return "selected" if they do
'*********************************************************************
function checkMatch(str1,str2)
	if lCase(trim(str1)) = lCase(trim(str2)) then
		checkMatch = " selected "
	else
		checkMatch = ""
	end if
end function
'*********************************************************************
'Display average rating for a product
'*********************************************************************
function ratingImage(prodRating)
	if not isNumeric(prodRating) then
		ratingImage = ""
		exit function
	end if
	select case round(prodRating,0)
	case 1
		ratingImage = "<img src=""../UserMods/misc_1rating.gif"" border=0 align=absmiddle>"
	case 2
		ratingImage = "<img src=""../UserMods/misc_2rating.gif"" border=0 align=absmiddle>"
	case 3
		ratingImage = "<img src=""../UserMods/misc_3rating.gif"" border=0 align=absmiddle>"
	case 4
		ratingImage = "<img src=""../UserMods/misc_4rating.gif"" border=0 align=absmiddle>"
	case 5
		ratingImage = "<img src=""../UserMods/misc_5rating.gif"" border=0 align=absmiddle>"
	case else
		ratingImage = ""
	end select
end function
'*********************************************************************
'Save a cart (order) for later retrieval
'*********************************************************************
function saveCart(idOrder,idCust)

	'Declare Variables
	dim mySQL, rsTemp, rsTemp2
	
	if isNumeric(idOrder) and isNumeric(idCust) then
	
		'Get some customer info
		mySQL="SELECT idCust,Name,LastName,CustomerCompany,Phone," _
			& "       Email,Address,City,Zip,locState,locCountry " _
		    & "FROM   customer " _
		    & "WHERE  idCust = " & validSQL(idCust,"I")
		set rsTemp = openRSexecute(mySQL)
		if not rstemp.eof then
		
			'Update cartHead
			mySQL = "UPDATE cartHead SET " _
				  & "orderStatus     = 'S'," _
				  & "idCust          = "  & validSQL(rsTemp("idCust"),"I")			& "," _
				  & "[Name]          = '" & validSQL(rsTemp("Name"),"A")			& "'," _
				  & "LastName        = '" & validSQL(rsTemp("LastName"),"A")		& "'," _
				  & "CustomerCompany = '" & validSQL(rsTemp("CustomerCompany"),"A")	& "'," _
				  & "Phone           = '" & validSQL(rsTemp("Phone"),"A")			& "'," _
				  & "Email           = '" & validSQL(rsTemp("Email"),"A")			& "'," _
				  & "Address         = '" & validSQL(rsTemp("Address"),"A")			& "'," _
				  & "City            = '" & validSQL(rsTemp("City"),"A")			& "'," _
				  & "Zip             = '" & validSQL(rsTemp("Zip"),"A")				& "'," _
				  & "locState        = '" & validSQL(rsTemp("locState"),"A")		& "'," _
				  & "locCountry      = '" & validSQL(rsTemp("locCountry"),"A")		& "' " _
				  & "WHERE idOrder   = "  & validSQL(idOrder,"I")
			set rsTemp2 = openRSexecute(mySQL)
			call closeRS(rsTemp2)

		end if
		call closeRS(rsTemp)

	end if

end function
'*************************************************************************
'Calculate an option's price for as it relates to a particular product.
'*************************************************************************
function getOptionPrice(priceToAdd, percToAdd, prodPrice)

	'Declare variables
	dim tempPrice

	'Check parameters
	if not(isNumeric(priceToAdd) and IsNumeric(percToAdd) and IsNumeric(prodPrice)) then
		getOptionPrice = 0
		exit function
	end if
	if isNull(priceToAdd) or isNull(percToAdd) or isNull(prodPrice) then
		getOptionPrice = 0
		exit function
	end if
	if priceToAdd = 0 and percToAdd = 0 then
		getOptionPrice = 0
		exit function
	end if

	'Determine Option Price
	if priceToAdd > 0 and percToAdd > 0 then
		tempPrice = Round(((prodPrice * percToAdd) / 100),2)
		if tempPrice > priceToAdd then
			getOptionPrice = tempPrice
		else
			getOptionPrice = priceToAdd
		end if
	elseif priceToAdd > 0 then
		getOptionPrice = priceToAdd
	else
		getOptionPrice = Round(((prodPrice * percToAdd) / 100),2)
	end if
	
end function
'*********************************************************************
'DEPRECATED Functions
'*********************************************************************
function checkString(str1)			'No longer required.
	checkString = str1
end function
function money(aNumber)				'Replaced by moneyS() and moneyD()
	money = moneyS(aNumber)
end function
%>
