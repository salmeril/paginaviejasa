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
'The purpose of the code in this script is to write an HTML form with 
'a "submit" button (or similar) to the browser. This form will "post" 
'some information to a payment gateway (the exact type of information 
'will depend on the gateway), allowing your customer to complete the 
'payment via the gateway of your choice. A good knowledge of the 
'gateway API will be required. 
'*************************************************************************
'
'--> YOUR CODE STARTS HERE
'
'This EXAMPLE creates a form for the LinkPoint Basic gateway. If you 
'want to use a gateway that is not pre-integrated, you will have to 
'replace the code below with code that is appropriate for your gateway.
'
'Enter your 'linkPointID' below.
const linkPointID = "000001"
'
%>
<form method="POST" action="https://secure.linkpt.net/cgi-bin/hlppay" name="CustomPaymentInfo">

	<!-- Standard LinkPoint fields -->
	<input type="hidden" name="chargetotal"	value="<%=moneyD(total)%>">
	<input type="hidden" name="mode"		value="payplus">
	<input type="hidden" name="storename"	value="<%=linkPointID%>">
	<input type="hidden" name="bname"		value="<%=name & " " & lastName%>">
	<input type="hidden" name="baddr1"		value="<%=address%>">
	<input type="hidden" name="bcity"		value="<%=city%>">
	<input type="hidden" name="bzip"		value="<%=zip%>">
	<input type="hidden" name="bcountry"	value="<%=countryCode%>">
	<input type="hidden" name="oid"			value="<%=pOrderPrefix & "-" & qIDOrder%>">
	
	<!-- LinkPoint has different fields for US and Other states -->
<%  if UCase(countryCode) = "US" then %>  
	<input type="hidden" name="bstate"		value="<%=stateCode%>">
<%	else  %>	
	<input type="hidden" name="bstate2"		value="<%=locState%>">
<%	end if%>
	
	<!-- Custom pass-through fields -->
	<input type="hidden" name="CP_idOrder"	value="<%=qIDOrder%>">
	
	<!-- Show message and submit button -->
	<br>
	<center>
		<b><font color=red>IMPORTANT&nbsp;:&nbsp;</font>Click below to submit payment.</b>
		<br><br>
		<b><font color=red size=2>--&gt;&nbsp;&nbsp;&nbsp;</font></b>
		<input type="submit" name="submit" value="Click to Pay">
		<b><font color=red size=2>&nbsp;&nbsp;&nbsp;&lt;--</font></b>
	</center>
	
</form>
<%
'--> YOUR CODE ENDS HERE
%>