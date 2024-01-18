<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Loads and Manages the Store Configuration settings
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
'Set session default LCID to 1033 - US English
'*************************************************************************
session.LCID = 1033

'*************************************************************************
'Declare local configuration variables
'*************************************************************************
dim urlNonSSL,urlSSL,pDownloadDir,pImagesDir,mailComp
dim pSMTPServer,pEmailSales,pEmailAdmin,pCompany,pCatalogOnly
dim pMaxCartQty,prodViewLayout,pHidePricingZero,pMaxItemQty
dim pMinCartAmount,pMaxItemsPerPage,pOrderPrefix,pCurrencySign
dim pStoreLCID,pShowStockView,pMailIn,pPayPal,payPalMemberID
dim pCreditCard,pCCType,pAuthNet,authNetLogin,authNetCurrCode
dim payMsgMailIn,payMsgCreditCard,payMsgPayPal,payMsgOther
dim payMsgNotReq,shipDisplayType,TwoCheckOut,TwoCheckOutSID
dim payMsgTwoCheckOut,pEmailFriendSec,pMaxDownloadHours
dim pMaxDownloadCount,payDefault,pAuthNetFrontEnd,payMsgAuthNet
dim pCompanyAddr,TwoCheckoutMD5,pHideAddStockLevel,payCustom
dim payMsgCustom,taxOnShipping,allowShipAddr,defaultCountryCode
dim listViewLayout,payCallIn,payFaxIn,PayCOD,statUpdPending
dim payMsgCallIn,payMsgFaxIn,payMsgCOD,taxBillOrShip
dim handlingFeeAmt,handlingFeeTax,payPalCurrCode,authNetTxKey

'*************************************************************************
'Get additional configuration settings from "config.asp"
'*************************************************************************
%><!--#include file="../config/config.asp"--><%

'*************************************************************************
'Retrieve configuration settings from DB and load into local variables
'*************************************************************************
function loadConfig()

	'Work variables
	dim mySQL, rsTemp
	dim configArr
	
	'Get configuration control record from database
	mySQL = "SELECT configValLong " _
		  & "FROM   storeAdmin " _
		  & "WHERE  configVar = 'controlRec' " _
		  & "AND    adminType = 'C' "
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.EOF then
		configArr = trim(rsTemp("configValLong"))
		if len(configArr) > 0 then
			configArr = split(configArr,"*|*")
		end if
	end if
	call closeRS(rsTemp)
	
	'Check config array
	if isNull(configArr) or isEmpty(configArr) or not(isArray(configArr)) then
		loadConfig = false
		exit function
	elseif UBound(configArr) <> 63 then
		loadConfig = false
		exit function
	end if

	'Assign config setting to local variables.
	on error resume next
	urlNonSSL			= Cstr(configArr(0))
	urlSSL				= Cstr(configArr(1))
	pDownloadDir		= Cstr(configArr(2))
	pImagesDir			= Cstr(configArr(3))
	mailComp			= CLng(configArr(4))
	pSMTPServer			= Cstr(configArr(5))
	pEmailSales			= Cstr(configArr(6))
	pEmailAdmin			= Cstr(configArr(7))
	pCompany			= Cstr(configArr(8))
	pCatalogOnly		= CLng(configArr(9))
	pMaxCartQty			= CLng(configArr(10))
	pMaxItemQty			= CLng(configArr(11))
	pMinCartAmount		= CLng(configArr(12))
	pMaxItemsPerPage	= CLng(configArr(13))
	pOrderPrefix		= Cstr(configArr(14))
	pCurrencySign		= Cstr(configArr(15))
	pStoreLCID			= Cstr(configArr(16))
	pShowStockView		= CLng(configArr(17))
	pMailIn				= CLng(configArr(18))
	pPayPal				= CLng(configArr(19))
	payPalMemberID		= Cstr(configArr(20))
	pCreditCard			= CLng(configArr(21))
	pCCType				= Cstr(configArr(22))
	pAuthNet			= CLng(configArr(23))
	authNetLogin		= Cstr(configArr(24))
	authNetCurrCode		= Cstr(configArr(25))
	payMsgMailIn		= Cstr(configArr(26))
	payMsgCreditCard	= Cstr(configArr(27))
	payMsgPayPal		= Cstr(configArr(28))
	payMsgOther			= Cstr(configArr(29))
	payMsgNotReq		= Cstr(configArr(30))
	TwoCheckOut			= CLng(configArr(31))
	TwoCheckOutSID		= Cstr(configArr(32))
	payMsgTwoCheckOut	= Cstr(configArr(33))
	pEmailFriendSec		= CLng(configArr(34))
	pMaxDownloadHours	= CLng(configArr(35))
	pMaxDownloadCount	= CLng(configArr(36))
	payDefault			= Cstr(configArr(37))
	pAuthNetFrontEnd	= CLng(configArr(38))
	pCompanyAddr		= Cstr(configArr(39))
	payMsgAuthNet		= Cstr(configArr(40))
	TwoCheckoutMD5		= Cstr(configArr(41))
	pHideAddStockLevel	= CLng(configArr(42))
	payCustom			= CLng(configArr(43))
	payMsgCustom		= Cstr(configArr(44))
	taxOnShipping		= CLng(configArr(45))
	allowShipAddr		= CLng(configArr(46))
	prodViewLayout		= CLng(configArr(47))
	shipDisplayType		= CLng(configArr(48))
	defaultCountryCode  = Cstr(configArr(49))
	pHidePricingZero	= CLng(configArr(50))
	payCallIn			= CLng(configArr(51))
	payFaxIn			= CLng(configArr(52))
	payCOD				= CLng(configArr(53))
	payMsgCallIn		= Cstr(configArr(54))
	payMsgFaxIn			= Cstr(configArr(55))
	payMsgCOD			= Cstr(configArr(56))
	listViewLayout		= CLng(configArr(57))
	taxBillOrShip		= CLng(configArr(58))
	statUpdPending		= CLng(configArr(59))
	handlingFeeAmt		= CDbl(configArr(60))
	handlingFeeTax		= CLng(configArr(61))
	payPalCurrCode		= Cstr(configArr(62))
	authNetTxKey		= Cstr(configArr(63))
	
	if err.number = 0 then
		loadConfig = true
	else
		loadConfig = false
	end if
	
	on error goto 0
	
end function
%>