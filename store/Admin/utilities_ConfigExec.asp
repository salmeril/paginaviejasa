<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Modify Store General Settings
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
<!--#include file="../Config/config.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<%
'Config Variables
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

'Control record
dim controlRec

'Database variables
dim mySQL, cn, rs

'Work variables
dim errCount
'*************************************************************************

'Are we in test mode?
if demoMode = "Y" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("DEMO MODE. Sorry, this featured is NOT available in Demo Mode.")
end if

'Initialize errCount
errCount = 0

'Get variables from the Form
urlNonSSL			= trim(Request.Form("urlNonSSL"))
urlSSL				= trim(Request.Form("urlSSL"))
pDownloadDir		= trim(Request.Form("pDownloadDir"))
pImagesDir			= trim(Request.Form("pImagesDir"))
mailComp			= trim(Request.Form("mailComp"))
pSMTPServer			= trim(Request.Form("pSMTPServer"))
pEmailSales			= trim(Request.Form("pEmailSales"))
pEmailAdmin			= trim(Request.Form("pEmailAdmin"))
pCompany			= trim(Request.Form("pCompany"))
pCatalogOnly		= trim(Request.Form("pCatalogOnly"))
pHidePricingZero	= trim(Request.Form("pHidePricingZero"))
pMaxCartQty			= trim(Request.Form("pMaxCartQty"))
pMaxItemQty			= trim(Request.Form("pMaxItemQty"))
pMinCartAmount		= trim(Request.Form("pMinCartAmount"))
pMaxItemsPerPage	= trim(Request.Form("pMaxItemsPerPage"))
pOrderPrefix		= trim(Request.Form("pOrderPrefix"))
pCurrencySign		= trim(Request.Form("pCurrencySign"))
pStoreLCID			= trim(Request.Form("pStoreLCID"))
pShowStockView		= trim(Request.Form("pShowStockView"))
pMailIn				= trim(Request.Form("pMailIn"))
pPayPal				= trim(Request.Form("pPayPal"))
payPalMemberID		= trim(Request.Form("payPalMemberID"))
TwoCheckOut			= trim(Request.Form("TwoCheckOut"))
TwoCheckOutSID		= trim(Request.Form("TwoCheckOutSID"))
pCreditCard			= trim(Request.Form("pCreditCard"))
pCCType				= trim(Request.Form("pCCType"))
pAuthNet			= trim(Request.Form("pAuthNet"))
authNetLogin		= trim(Request.Form("authNetLogin"))
authNetCurrCode		= trim(Request.Form("authNetCurrCode"))
payMsgMailIn		= trim(Request.Form("payMsgMailIn"))
payMsgCreditCard	= trim(Request.Form("payMsgCreditCard"))
payMsgPayPal		= trim(Request.Form("payMsgPayPal"))
payMsgTwoCheckOut	= trim(Request.Form("payMsgTwoCheckOut"))
payMsgOther			= trim(Request.Form("payMsgOther"))
payMsgNotReq		= trim(Request.Form("payMsgNotReq"))
pEmailFriendSec		= trim(Request.Form("pEmailFriendSec"))
pMaxDownloadHours	= trim(Request.Form("pMaxDownloadHours"))
pMaxDownloadCount	= trim(Request.Form("pMaxDownloadCount"))
payDefault			= trim(Request.Form("payDefault"))
pAuthNetFrontEnd	= trim(Request.Form("pAuthNetFrontEnd"))
pCompanyAddr		= trim(Request.Form("pCompanyAddr"))
payMsgAuthNet		= trim(Request.Form("payMsgAuthNet"))
TwoCheckoutMD5		= trim(Request.Form("TwoCheckoutMD5"))
pHideAddStockLevel	= trim(Request.Form("pHideAddStockLevel"))
payCustom			= trim(Request.Form("payCustom"))
payMsgCustom		= trim(Request.Form("payMsgCustom"))
taxOnShipping		= trim(Request.Form("taxOnShipping"))
allowShipAddr		= trim(Request.Form("allowShipAddr"))
prodViewLayout		= trim(Request.Form("prodViewLayout"))
shipDisplayType		= trim(Request.Form("shipDisplayType"))
defaultCountryCode	= trim(Request.Form("defaultCountryCode"))
payCallIn			= trim(Request.Form("payCallIn"))
payFaxIn			= trim(Request.Form("payFaxIn"))
payCOD				= trim(Request.Form("payCOD"))
payMsgCallIn		= trim(Request.Form("payMsgCallIn"))
payMsgFaxIn			= trim(Request.Form("payMsgFaxIn"))
payMsgCOD			= trim(Request.Form("payMsgCOD"))
listViewLayout		= trim(Request.Form("listViewLayout"))
taxBillOrShip		= trim(Request.Form("taxBillOrShip"))
statUpdPending		= trim(Request.Form("statUpdPending"))
handlingFeeAmt		= trim(Request.Form("handlingFeeAmt"))
handlingFeeTax		= trim(Request.Form("handlingFeeTax"))
payPalCurrCode		= trim(Request.Form("payPalCurrCode"))
authNetTxKey		= trim(Request.Form("authNetTxKey"))
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Store Configuration</font></b>
</P>

<table border=0 cellspacing=0 cellpadding=10 width="100%" class="textBlock">
<tr><td>
<%

'urlNonSSL
if not (len(urlNonSSL) > 0 and (left(lCase(urlNonSSL),7)="http://" or left(lCase(urlNonSSL),8)="https://") and right(urlNonSSL,1)="/") then
	Response.Write showTest("urlNonSSL")
end if

'urlSSL
if not (len(urlSSL) > 0 and (left(lCase(urlSSL),7)="http://" or left(lCase(urlSSL),8)="https://") and right(urlSSL,1)="/") then
	Response.Write showTest("urlSSL")
end if

'pDownloadDir
if not (len(pDownloadDir) > 0 and left(pDownloadDir,3)="../" and right(pDownloadDir,1)="/") then
	Response.Write showTest("pDownloadDir")
end if

'pImagesDir
if not (len(pImagesDir) > 0 and left(pImagesDir,3)="../" and right(pImagesDir,1)="/") then
	Response.Write showTest("pImagesDir")
end if

'mailComp
if  mailComp <> "0" _
and mailComp <> "1" _
and mailComp <> "2" _
and mailComp <> "3" _
and mailComp <> "4" _
and mailComp <> "5" _
and mailComp <> "6" then
	Response.Write showTest("mailComp")
end if

'pSMTPServer
if mailComp <> "0" then
	if len(pSMTPServer) = 0 then
		Response.Write showTest("pSMTPServer")
	end if
end if

'pEmailSales
if not (len(pEmailSales) > 0 and instr(pEmailSales,"@") > 0 and instr(pEmailSales,".") > 0) then
	Response.Write showTest("pEmailSales")
end if

'pEmailAdmin
if not (len(pEmailAdmin) > 0 and instr(pEmailAdmin,"@") > 0 and instr(pEmailAdmin,".") > 0) then
	Response.Write showTest("pEmailAdmin")
end if
	
'pCompany
if len(pCompany) = 0 then
	Response.Write showTest("pCompany")
end if

'pCompanyAddr
if len(pCompanyAddr) > 250 then
	Response.Write showTest("pCompanyAddr")
end if

'pCatalogOnly
if not (pCatalogOnly="0" or pCatalogOnly="-1") then
	Response.Write showTest("pCatalogOnly")
end if

'prodViewLayout
if  prodViewLayout <> "0" _
and prodViewLayout <> "1" then
	Response.Write showTest("prodViewLayout")
end if

'listViewLayout
if  listViewLayout <> "0" _
and listViewLayout <> "1" then
	Response.Write showTest("listViewLayout")
end if

'shipDisplayType
if  shipDisplayType <> "0" _
and shipDisplayType <> "1" then
	Response.Write showTest("shipDisplayType")
end if

'allowShipAddr
if not (allowShipAddr="0" or allowShipAddr="-1") then
	Response.Write showTest("allowShipAddr")
end if

'taxOnShipping
if not (taxOnShipping="0" or taxOnShipping="-1") then
	Response.Write showTest("taxOnShipping")
end if

'handlingFeeAmt
if not isnumeric(handlingFeeAmt) then
	Response.Write showTest("handlingFeeAmt")
end if

'handlingFeeTax
if not (handlingFeeTax="0" or handlingFeeTax="-1") then
	Response.Write showTest("handlingFeeTax")
end if

'taxBillOrShip
if  taxBillOrShip <> "0" _
and taxBillOrShip <> "1" then
	Response.Write showTest("taxBillOrShip")
end if

'pHidePricingZero
if not (pHidePricingZero="0" or pHidePricingZero="-1") then
	Response.Write showTest("pHidePricingZero")
end if

'pShowStockView
if not (pShowStockView="0" or pShowStockView="-1") then
	Response.Write showTest("pShowStockView")
end if

'pHideAddStockLevel
if not isnumeric(pHideAddStockLevel) then
	Response.Write showTest("pHideAddStockLevel")
end if

'statUpdPending
if not (statUpdPending="0" or statUpdPending="-1") then
	Response.Write showTest("statUpdPending")
end if

'pMaxCartQty
if not isnumeric(pMaxCartQty) then
	Response.Write showTest("pMaxCartQty")
end if

'pMaxItemQty
if not isnumeric(pMaxItemQty) then
	Response.Write showTest("pMaxItemQty")
end if

'pMinCartAmount
if not isnumeric(pMinCartAmount) then
	Response.Write showTest("pMinCartAmount")
end if

'pMaxItemsPerPage
if not isnumeric(pMaxItemsPerPage) then
	Response.Write showTest("pMaxItemsPerPage")
end if

'pOrderPrefix
if len(pOrderPrefix) = 0 then
	Response.Write showTest("pOrderPrefix")
end if

'pCurrencySign
if len(pCurrencySign) = 0 then
	Response.Write showTest("pCurrencySign")
end if

'pStoreLCID
on error resume next
session.LCID = pStoreLCID
if err.number <> 0 then
	Response.Write showTest("pStoreLCID")
end if
err.Clear
on error goto 0
session.LCID = 1033

'pEmailFriendSec
if not (pEmailFriendSec="0" or pEmailFriendSec="-1") then
	Response.Write showTest("pEmailFriendSec")
end if

'pMaxDownloadHours
if not isnumeric(pMaxDownloadHours) then
	Response.Write showTest("pMaxDownloadHours")
end if

'pMaxDownloadCount
if not isnumeric(pMaxDownloadCount) then
	Response.Write showTest("pMaxDownloadCount")
end if

'pMailIn
if not (pMailIn="0" or pMailIn="-1") then
	Response.Write showTest("pMailIn")
end if

'payMsgMailIn
if len(payMsgMailIn) = 0 then
	Response.Write showTest("payMsgMailIn")
end if

'payCallIn
if not (payCallIn="0" or payCallIn="-1") then
	Response.Write showTest("payCallIn")
end if

'payMsgCallIn
if len(payMsgCallIn) = 0 then
	Response.Write showTest("payMsgCallIn")
end if

'payFaxIn
if not (payFaxIn="0" or payFaxIn="-1") then
	Response.Write showTest("payFaxIn")
end if

'payMsgFaxIn
if len(payMsgFaxIn) = 0 then
	Response.Write showTest("payMsgFaxIn")
end if

'payCOD
if not (payCOD="0" or payCOD="-1") then
	Response.Write showTest("payCOD")
end if

'payMsgCOD
if len(payMsgCOD) = 0 then
	Response.Write showTest("payMsgCOD")
end if

'pPayPal
if not (pPayPal="0" or pPayPal="-1") then
	Response.Write showTest("pPayPal")
end if

'payPalMemberID
if pPayPal="-1" then
	if len(payPalMemberID) = 0 then
		Response.Write showTest("payPalMemberID")
	end if
end if

'payMsgPayPal
if len(payMsgPayPal) = 0 then
	Response.Write showTest("payMsgPayPal")
end if

'payPalCurrCode
if len(payPalCurrCode) <> 3 then
	Response.Write showTest("payPalCurrCode")
end if

'TwoCheckOut
if not (TwoCheckOut="0" or TwoCheckOut="-1") then
	Response.Write showTest("TwoCheckOut")
end if

'TwoCheckOutSID
if TwoCheckOut="-1" then
	if len(TwoCheckOutSID) = 0 then
		Response.Write showTest("TwoCheckOutSID")
	end if
end if

'payMsgTwoCheckOut
if len(payMsgTwoCheckOut) = 0 then
	Response.Write showTest("payMsgTwoCheckOut")
end if

'pCreditCard
if not (pCreditCard="0" or pCreditCard="-1") then
	Response.Write showTest("pCreditCard")
end if

'pCCType
if pCreditCard="-1" then
	if len(pCCType) = 0 then
		Response.Write showTest("pCCType")
	end if
end if

'payMsgCreditCard
if len(payMsgCreditCard) = 0 then
	Response.Write showTest("payMsgCreditCard")
end if

'pAuthNetFrontEnd
if not (pAuthNetFrontEnd="0" or pAuthNetFrontEnd="-1") then
	Response.Write showTest("pAuthNetFrontEnd")
end if

'pAuthNet
if not (pAuthNet="0" or pAuthNet="-1") then
	Response.Write showTest("pAuthNet")
end if

'authNetLogin & authNetCurrCode & authNetTxKey
if pAuthNet="-1" or pAuthNetFrontEnd="-1" then
	if len(authNetLogin) = 0 then
		Response.Write showTest("authNetLogin")
	end if
	if len(authNetCurrCode) = 0 then
		Response.Write showTest("authNetCurrCode")
	end if
	if len(authNetTxKey) = 0 then
		Response.Write showTest("authNetTxKey")
	end if
end if

'payMsgAuthNet
if len(payMsgAuthNet) = 0 then
	Response.Write showTest("payMsgAuthNet")
end if

'payCustom
if not (payCustom="0" or payCustom="-1") then
	Response.Write showTest("payCustom")
end if

'payMsgCustom
if len(payMsgCustom) = 0 then
	Response.Write showTest("payMsgCustom")
end if

'payMsgOther
if len(payMsgOther) = 0 then
	Response.Write showTest("payMsgOther")
end if

'payMsgNotReq
if len(payMsgNotReq) = 0 then
	Response.Write showTest("payMsgNotReq")
end if

'Control record
controlRec = "" _	
	& urlNonSSL				& "*|*" _
	& urlSSL				& "*|*" _
	& pDownloadDir			& "*|*" _
	& pImagesDir			& "*|*" _
	& mailComp				& "*|*" _
	& pSMTPServer			& "*|*" _
	& pEmailSales			& "*|*" _
	& pEmailAdmin			& "*|*" _
	& pCompany				& "*|*" _
	& pCatalogOnly			& "*|*" _
	& pMaxCartQty			& "*|*" _
	& pMaxItemQty			& "*|*" _
	& pMinCartAmount		& "*|*" _
	& pMaxItemsPerPage		& "*|*" _
	& pOrderPrefix			& "*|*" _
	& pCurrencySign			& "*|*" _
	& pStoreLCID			& "*|*" _
	& pShowStockView		& "*|*" _
	& pMailIn				& "*|*" _
	& pPayPal				& "*|*" _
	& payPalMemberID		& "*|*" _
	& pCreditCard			& "*|*" _
	& pCCType				& "*|*" _
	& pAuthNet				& "*|*" _
	& authNetLogin			& "*|*" _
	& authNetCurrCode		& "*|*" _
	& payMsgMailIn			& "*|*" _
	& payMsgCreditCard		& "*|*" _
	& payMsgPayPal			& "*|*" _
	& payMsgOther			& "*|*" _
	& payMsgNotReq			& "*|*" _
	& TwoCheckOut			& "*|*" _
	& TwoCheckOutSID		& "*|*" _
	& payMsgTwoCheckOut		& "*|*" _
	& pEmailFriendSec		& "*|*" _
	& pMaxDownloadHours		& "*|*" _
	& pMaxDownloadCount		& "*|*" _
	& payDefault			& "*|*" _
	& pAuthNetFrontEnd		& "*|*" _
	& pCompanyAddr			& "*|*" _
	& payMsgAuthNet			& "*|*" _
	& TwoCheckoutMD5		& "*|*" _
	& pHideAddStockLevel	& "*|*" _
	& payCustom				& "*|*" _
	& payMsgCustom			& "*|*" _
	& taxOnShipping			& "*|*" _
	& allowShipAddr			& "*|*" _
	& prodViewLayout		& "*|*" _
	& shipDisplayType		& "*|*" _
	& defaultCountryCode	& "*|*" _
	& pHidePricingZero		& "*|*" _
	& payCallIn				& "*|*" _
	& payFaxIn				& "*|*" _
	& payCOD				& "*|*" _
	& payMsgCallIn			& "*|*" _
	& payMsgFaxIn			& "*|*" _
	& payMsgCOD				& "*|*" _
	& listViewLayout		& "*|*" _
	& taxBillOrShip			& "*|*" _
	& statUpdPending		& "*|*" _
	& handlingFeeAmt		& "*|*" _
	& handlingFeeTax		& "*|*" _
	& payPalCurrCode		& "*|*" _
	& authNetTxKey
	
'Check that there are no double quotes (") in any of the fields
if instr(controlRec,"""") > 0 then
	Response.Write showTest("Double Quotes are not allowed in any of the configuration settings.")
end if

'Check if any errors were detected
if errCount > 0 then
%>
	<br>
	<b><font color=red><%=errCount%> Errors were detected</font></b>
<%
else

	'Update the configuration file
	call writeConfigFile()
%>
	<br>
	<b><font size=3 color=green>Success!</font></b><br><br><br>
	<font size=2>
		Store Configurations were successfully updated.
	</font><br><br>
<%
end if
%>
</td></tr>
</table>
<!--#include file="_INCfooter_.asp"-->
<%
'**********************************************************************
'Display Test / Result
'**********************************************************************
function showTest(testStr)
	showTest = ""
	errCount = errCount + 1		
	if errCount = 1 then
		showTest = "<br><b>Some errors were detected...</b><br><br>"
	end if
	showTest = showTest _
		& "<img src=x_N.gif border=0 valign=absMiddle> " _ 
		& testStr _
		& "<br>"
end function
'**********************************************************************
'Update the Configurations
'**********************************************************************
sub writeConfigFile()

	'Create an array of config variable names to drive the DB update
	dim configArr, I
	configArr = "" _
		& "urlNonSSL,urlSSL,pDownloadDir,pImagesDir,mailComp," _
		& "pSMTPServer,pEmailSales,pEmailAdmin,pCompany,pCatalogOnly," _
		& "pMaxCartQty,pMaxItemQty,pMinCartAmount,pMaxItemsPerPage," _
		& "pOrderPrefix,pCurrencySign,pStoreLCID,pShowStockView," _
		& "pMailIn,pPayPal,payPalMemberID,pCreditCard,pCCType," _
		& "pAuthNet,authNetLogin,authNetCurrCode,payMsgMailIn," _
		& "payMsgCreditCard,payMsgPayPal,payMsgOther,payMsgNotReq," _
		& "TwoCheckOut,TwoCheckOutSID,payMsgTwoCheckOut," _
		& "pEmailFriendSec,pMaxDownloadHours,pMaxDownloadCount,payDefault," _
		& "pAuthNetFrontEnd,pCompanyAddr,payMsgAuthNet,TwoCheckoutMD5," _
		& "pHideAddStockLevel,payCustom,payMsgCustom,taxOnShipping," _
		& "allowShipAddr,prodViewLayout,shipDisplayType," _
		& "defaultCountryCode,pHidePricingZero,statUpdPending," _
		& "payCallIn,payFaxIn,payCOD,payMsgCallIn,payMsgFaxIn," _
		& "payMsgCOD,listViewLayout,taxBillOrShip,handlingFeeAmt," _
		& "handlingFeeTax,payPalCurrCode,authNetTxKey"
	configArr = split(configArr,",")
	
	'Open Database
	call openDb()

	'Loop through array and UPDATE / INSERT new config settings
	for I = 0 to UBound(configArr)
		if len(trim(configArr(I))) > 0 then
			mySQL = "SELECT configVal " _
				&   "FROM   storeAdmin " _
				&   "WHERE  configVar='" & configArr(I) & "' " _
				&   "AND    adminType='C'"
			set rs = openRSexecute(mySQL)
			if rs.EOF then
				'INSERT
				call closeRS(rs)
				mySQL = "INSERT INTO storeAdmin " _
					  & "(adminType,configVar,configVal) " _
					  & "VALUES " _
					  & "('C','" & configArr(I) & "','" & replace(eval(configArr(I)),"'","''") & "')"
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
						  & "AND    adminType='C'"
					set rs = openRSexecute(mySQL)
				end if
			end if
		end if
	next
	
	'UPDATE / INSERT control record
	mySQL = "SELECT configVar " _
		&   "FROM   storeAdmin " _
		&   "WHERE  configVar='controlRec' " _
		&   "AND    adminType='C'"
	set rs = openRSexecute(mySQL)
	if rs.EOF then
		'INSERT
		call closeRS(rs)
		mySQL = "INSERT INTO storeAdmin " _
			  & "(adminType,configVar,configValLong) " _
			  & "VALUES " _
			  & "('C','controlRec','" & replace(controlRec,"'","''") & "')"
		set rs = openRSexecute(mySQL)
	else
		'UPDATE
		call closeRS(rs)
		mySQL = "UPDATE storeAdmin SET " _
		      & "       configValLong='" & replace(controlRec,"'","''") & "' " _
			  & "WHERE  configVar='controlRec' " _
			  & "AND    adminType='C'"
		set rs = openRSexecute(mySQL)
	end if

	'Close Database
	call closedb()
	
end sub

%>