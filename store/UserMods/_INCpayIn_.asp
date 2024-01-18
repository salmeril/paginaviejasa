<%
'*************************************************************************
' Function : Userexit that can be used to create a customized 
'          : payment routine. Should only be modified by someone
'          : with experience in ASP development.
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
'OVERVIEW :
'The purpose of the code in this script is to populate 2 variables when 
'the customer is returned to your store after submitting payment through 
'a custom payment gateway. This process will vary from one payment 
'processor to the next.
' 
'qIdOrder  = Order Number (Numeric)
'statusInd = Must contain "error" or "success"
'
'*************************************************************************
'--> YOUR CODE STARTS HERE
'
'This EXAMPLE captures the return values for the LinkPoint Basic gateway. 
'If you want to use a gateway that is not pre-integrated, you will have 
'to replace the code below with code that is appropriate for your gateway.
'
if len(qIdOrder) = 0 then

	'Get Order Number (try the Form first, then the QueryString)
	qIdOrder = trim(Request.Form("CP_idOrder"))
	if len(qIdOrder) = 0 then
		qIdOrder = trim(Request.QueryString("CP_idOrder"))
	end if
	
	'If the above lines didn't work, let's try another way
	if len(qIdOrder) = 0 then
		qIdOrder = trim(Request.Form("oid"))
		if len(qIdOrder) > 0 then
			qIdOrder = replace(qIdOrder,pOrderPrefix&"-","")
		end if
	end if
	
	'Get Status
	if len(qIdOrder) > 0 then
		statusInd = trim(Request.Form("status"))
		if lCase(statusInd) = "approved" then
			statusInd = "success"
		else
			statusInd = "error"
		end if
	end if
	
end if
'
'--> YOUR CODE ENDS HERE
%>