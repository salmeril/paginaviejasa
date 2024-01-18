<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Security & User Authentication
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
dim adminLoggedOn

'Get logon level from session
adminLoggedOn = session(storeID & "adminLoggedOn")

'Check if there is a valid logon
if isEmpty(adminLoggedOn) or isNull(adminLoggedOn) or not isNumeric(adminLoggedOn) then
	session.abandon 
	response.redirect "logon.asp" 
else
	adminLoggedOn = CLng(adminLoggedOn)
	if adminLoggedOn <> 0 and adminLoggedOn <> 1 then
		session.abandon 
		response.redirect "logon.asp" 
	end if
end if

'Check if logon is allowed to access the page
if adminLevel < adminLoggedOn then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("You are not authorized to view this page.")
end if
%>