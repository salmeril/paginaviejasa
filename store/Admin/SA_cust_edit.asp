<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Customer Maintenance
' Product  : CandyPress eCommerce Administration
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
const adminLevel = 1
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="../Scripts/_INCappFunctions_.asp"-->
<!--#include file="../Scripts/_INCrc4_.asp"-->
<%

'Database
dim mySQL, cn, rs, rs2

'Customer
dim idCust
dim status
dim dateCreated
dim dateCreatedInt
dim name
dim lastName
dim customerCompany
dim phone
dim email
dim password
dim address
dim city
dim locState
dim locState2
dim locCountry
dim zip
dim paymentType
dim shippingName
dim shippingLastName
dim shippingPhone
dim shippingAddress
dim shippingCity
dim shippingLocState
dim shippingLocState2
dim shippingLocCountry
dim shippingZip
dim futureMail
dim generalComments
dim taxExempt

'Work Fields
dim action
dim orderCount

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Customer Maintenance</font></b>
	<br><br>
</P>

<%
'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if action <> "edit" and action <> "del" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idCust
idCust = trim(Request.QueryString("recId"))
if len(idCust) = 0 then
	idCust = trim(Request.Form("recId"))
end if
if idCust = "" or not isNumeric(idCust) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
end if

'Get customer record
mySQL	= "SELECT status,dateCreated,password,name,lastName," _
		& "       customerCompany,phone,email,address,city," _
		& "       locState,locState2,locCountry,zip,shippingName," _
		& "       shippingLastName,shippingPhone,shippingAddress,"_
		& "       shippingCity,shippingLocState,shippingLocState2," _
		& "       shippingLocCountry,shippingZip,paymentType," _
		& "       futureMail,taxExempt,generalComments " _
		& "FROM   customer " _
		& "WHERE  idCust=" & idCust
set rs = openRSexecute(mySQL)
if rs.eof then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
else
	generalComments		= trim(rs("generalComments"))
	status				= trim(rs("status"))
	dateCreated			= rs("dateCreated")
	password			= trim(EnDeCrypt(Hex2Ascii(rs("password")),rc4Key))
	Name				= trim(rs("name"))
	LastName			= trim(rs("LastName"))
	CustomerCompany		= trim(rs("CustomerCompany"))
	Phone				= trim(rs("Phone"))
	Email				= trim(rs("Email"))
	Address				= trim(rs("Address"))
	City				= trim(rs("City"))
	locState			= trim(rs("locState"))
	locState2			= trim(rs("locState2"))
	locCountry			= trim(rs("locCountry"))
	Zip					= trim(rs("Zip"))
	shippingName		= trim(rs("shippingName"))
	shippingLastName	= trim(rs("shippingLastName"))
	shippingPhone		= trim(rs("shippingPhone"))
	shippingAddress		= trim(rs("shippingAddress"))
	ShippingCity		= trim(rs("ShippingCity"))
	shippingLocState	= trim(rs("shippingLocState"))
	shippingLocState2	= trim(rs("shippingLocState2"))
	shippingLocCountry	= trim(rs("shippingLocCountry"))
	shippingZip			= trim(rs("shippingZip"))
	paymentType			= trim(rs("paymentType"))
	futureMail			= trim(rs("futureMail"))
	taxExempt			= trim(rs("taxExempt"))
end if
call closeRS(rs)

'Get number of Orders for this Customer
mySQL = "SELECT COUNT(*) AS orderCount " _
	  & "FROM   cartHead " _
	  & "WHERE  idCust=" & idCust & " "
set rs = openRSexecute(mySQL)
orderCount = rs("orderCount")
call closeRS(rs)

if len(trim(Request.QueryString("msg"))) > 0 then
%>
	<font color=red><%=Request.QueryString("msg")%></font>
	<br><br>
<%
end if

if action = "del" then
%>
	<span class="textBlockHead">Delete Customer</span>
	&nbsp;<%call maintNavLinks()%><br>
	
	<form method="post" action="SA_cust_exec.asp" name="form4">
		<font color=red>Are you sure you want to Delete this Customer?</font>
		<input type=hidden name=idCust  id=idCust  value="<%=idCust%>">
		<input type=hidden name=action  id=action  value="del">
		<input type=submit name=submit1 id=submit1 value="Yes">
	</form>
	
	<table border=0 cellspacing=0 cellpadding=5 width="400" class="textBlock">
		<tr>
			<td bgcolor="#DDDDDD"><i>General</i></td>
			<td bgcolor="#DDDDDD" align="right">
				<a href="email.asp?emailTo=<%=server.URLEncode(email)%>&emailToName=<%=server.URLEncode(name & " " & LastName)%>">Send Email</a> 
			</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Status</b></td>
			<td align=left><%=status%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Customer ID</b></td>
			<td align=left><%=idCust%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Date Created</b></td>
			<td align=left><%=formatTheDate(dateCreated)%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Tax Exempt?</b></td>
			<td align=left><%=emptyString(taxExempt,"N")%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Number of Orders</b></td>
			<td align=left nowrap>
				<%=orderCount%>&nbsp;&nbsp;
				(<a href="SA_order.asp?showField=idCust&showPhrase=<%=idCust%>">List Orders</a>)
			</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD"><i>Customer Info</i></td>
		</tr>
		<tr>
			<td align=left nowrap><b>Name</b></td>
			<td align=left><%=lastName & ", " & name%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Company</b></td>
			<td align=left><%=customerCompany%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Phone #</b></td>
			<td align=left><%=phone%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Email</b></td>
			<td align=left><%=email%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Password</b></td>
			<td align=left><%=password%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Mailing List?</b></td>
			<td align=left><%=futureMail%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Payment Type</b></td>
			<td align=left><%=paymentType%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD"><i>Billing Address</i></td>
		</tr>
		<tr>
			<td align=left nowrap><b>Address </b></td>
			<td align=left><%=address%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>City </b></td>
			<td align=left><%=city%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Zip/Postal Code </b></td>
			<td align=left><%=zip%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>State/Province </b></td>
			<td align=left><%=getStateDesc(locCountry,locState,locState2)%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Country </b></td>
			<td align=left><%=getCountryDesc(locCountry)%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD"><i>Shipping Adrress</i></td>
		</tr>
		<tr>
			<td align=left nowrap><b>Name </b></td>
			<td align=left>
<%
				if len(shippingName) > 0 then
					Response.Write ShippingLastName & ", " & shippingName
				else
					Response.Write ShippingLastName
				end if
%>
				&nbsp;
			</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Phone #</b></td>
			<td align=left><%=shippingPhone%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Address </b></td>
			<td align=left><%=Shippingaddress%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>City </b></td>
			<td align=left><%=Shippingcity%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Zip/Postal Code </b></td>
			<td align=left><%=Shippingzip%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>State/Province </b></td>
			<td align=left><%=getStateDesc(shippingLocCountry,shippingLocState,shippingLocState2)%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Country </b></td>
			<td align=left><%=getCountryDesc(shippingLocCountry)%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD"><i><b>Comments</b> (NOT viewable by Customer)</i></td>
		</tr>
		<tr>
			<td align=left valign=top nowrap><b>Comments</b></td>
			<td align=left valign=top><%=generalComments%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD">&nbsp;</td>
		</tr>
	</table>
<%
end if

if action = "edit" then
%>
	<span class="textBlockHead">Edit Customer</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 width="400" class="textBlock">
		<form method="post" action="SA_cust_exec.asp" name="form1">
		<tr>
			<td bgcolor="#DDDDDD"><i>General</i></td>
			<td bgcolor="#DDDDDD" align="right">
				<a href="email.asp?emailTo=<%=server.URLEncode(email)%>&emailToName=<%=server.URLEncode(name & " " & LastName)%>">Send Email</a> 
			</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Status</b></td>
			<td align=left>
				<select name=status id=status size=1>
					<option value="A" <%=checkMatch(status,"A")%>>Active</option>
					<option value="I" <%=checkMatch(status,"I")%>>InActive</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Customer ID</b></td>
			<td align=left><%=idCust%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Date Created</b></td>
			<td align=left><%=formatTheDate(dateCreated)%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Tax Exempt?</b></td>
			<td align=left>
				<select name=taxExempt id=taxExempt size=1>
					<option value="N" <%=checkMatch(taxExempt,"N")%>>No</option>
					<option value="Y" <%=checkMatch(taxExempt,"Y")%>>Yes</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Number of Orders</b></td>
			<td align=left nowrap>
				<%=orderCount%>&nbsp;&nbsp;
				(<a href="SA_order.asp?showField=idCust&showPhrase=<%=idCust%>">List Orders</a>)
			</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD"><i>Customer Info</i></td>
		</tr>
		<tr>
			<td align=left nowrap><b>Name</b></td>
			<td align=left><%=lastName & ", " & name%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Company</b></td>
			<td align=left><%=customerCompany%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Phone #</b></td>
			<td align=left><%=phone%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Email</b></td>
			<td align=left>
				<input type=text name=email size=30 maxlength="50" value="<%=email%>">
			</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Password</b></td>
			<td align=left><%=password%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Mailing List?</b></td>
			<td align=left><%=futureMail%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Payment Type</b></td>
			<td align=left><%=paymentType%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD"><i>Billing Address</i></td>
		</tr>
		<tr>
			<td align=left nowrap><b>Address </b></td>
			<td align=left><%=address%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>City </b></td>
			<td align=left><%=city%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Zip/Postal Code </b></td>
			<td align=left><%=zip%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>State/Province </b></td>
			<td align=left><%=getStateDesc(locCountry,locState,locState2)%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Country </b></td>
			<td align=left><%=getCountryDesc(locCountry)%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD"><i>Shipping Adrress</i></td>
		</tr>
		<tr>
			<td align=left nowrap><b>Name </b></td>
			<td align=left>
<%
				if len(shippingName) > 0 then
					Response.Write ShippingLastName & ", " & shippingName
				else
					Response.Write ShippingLastName
				end if
%>
				&nbsp;
			</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Phone #</b></td>
			<td align=left><%=shippingPhone%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Address </b></td>
			<td align=left><%=Shippingaddress%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>City </b></td>
			<td align=left><%=Shippingcity%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Zip/Postal Code </b></td>
			<td align=left><%=Shippingzip%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>State/Province </b></td>
			<td align=left><%=getStateDesc(shippingLocCountry,shippingLocState,shippingLocState2)%>&nbsp;</td>
		</tr>
		<tr>
			<td align=left nowrap><b>Country </b></td>
			<td align=left><%=getCountryDesc(shippingLocCountry)%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan=2 bgcolor="#DDDDDD"><i><b>Comments</b> (NOT viewable by Customer)</i></td>
		</tr>
		<tr>
			<td colspan=2 align=center valign=middle>
				<textarea name=generalComments cols=45 rows=6><%=generalComments%></textarea>
			</td>
		</tr>
		<tr>
			<td colspan=2 align=center>
				<input type=hidden name=idCust  id=idCust  value="<%=idCust%>">
				<input type=hidden name=action  id=action  value="edit">
				<input type=submit name=submit1 id=submit1 value="Update Customer">
			</td>
		</tr>
		</form>
	</table>
<%
end if

if action = "edit" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>
	
		<b>Customer Status</b> - To prevent a Customer from logging on to 
		their account and viewing their Orders, set the Status to InActive.<br><br>
	
		<b>Tax Exempt</b> - If set to "Yes", taxes will NOT be calculated 
		on orders placed by this customer. If set to "No", taxes 
		will be calculated as normal.<br><br>
	
		<b>Email</b> - You may modify this field if required. Changing the 
		email here will not change the email address on the Customer's 
		orders.<br><br>
	
		<b>Comments</b> - These are private "notes" on the Customer. Any 
		text entered into this field can NOT be viewed by the Customer.<br><br>
	
	</td></tr>
	</table>
<%
end if

call closedb()
%>
<!--#include file="_INCfooter_.asp"-->
<%
'*********************************************************************
'Create Navigation Links
'*********************************************************************
sub maintNavLinks()
%>
	[ 
	<a href="SA_cust.asp?recallCookie=1">List Customers</a> | 
	<a href="SA_cust_edit.asp?action=edit&recid=<%=idCust%>">Edit Customer</a> | 
	<a href="SA_cust_edit.asp?action=del&recid=<%=idCust%>">Delete Customer</a> | 
	<a href="SA_order.asp?showField=idCust&showPhrase=<%=idCust%>">List Orders</a> 
	]
<%
end sub
%>
