<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Order Discount Maintenance
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
<%
'Database
dim mySQL, cn, rs

'DiscOrder
dim idDiscOrder
dim discCode
dim discPerc
dim discAmt
dim discFromAmt
dim discToAmt
dim discStatus
dim discOnceOnly
dim discValidFrom
dim discValidTo

'Work Fields
dim I
dim action

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
	<b><font size=3>Order Discount Maintenance</font></b>
	<br><br>
</P>

<%
'Get action
action = trim(Request.QueryString("action"))
if len(action) = 0 then
	action = trim(Request.Form("action"))
end if
action = lCase(action)
if action <> "edit" and action <> "add" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action Indicator.")
end if

'Get idDiscOrder
if action = "edit" then
	idDiscOrder = trim(Request.QueryString("recId"))
	if len(idDiscOrder) = 0 then
		idDiscOrder = trim(Request.Form("recId"))
	end if
	if idDiscOrder = "" or not isNumeric(idDiscOrder) then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	end if
end if

'Get Order Discount Record
if action = "edit" then
	mySQL="SELECT idDiscOrder,discCode,discPerc,discAmt," _
		& "       discFromAmt,discToAmt,discStatus," _
	    & "       discOnceOnly,discValidFrom,discValidTo " _
	    & "FROM   DiscOrder " _
	    & "WHERE  idDiscOrder = " & idDiscOrder
	set rs = openRSexecute(mySQL)
	if rs.eof then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Record ID.")
	else
		discCode		= rs("discCode")
		discPerc		= rs("discPerc")
		discAmt			= rs("discAmt")
		discFromAmt		= rs("discFromAmt")
		discToAmt		= rs("discToAmt")
		discStatus		= rs("discStatus")
		discOnceOnly	= rs("discOnceOnly")
		discValidFrom	= rs("discValidFrom")
		discValidTo		= rs("discValidTo")
	end if
	call closeRS(rs)
end if

'Edit
if action = "edit" then
	if len(trim(Request.QueryString("msg"))) > 0 then
%>
		<font color=red><%=Request.QueryString("msg")%></font>
		<br><br>
<%
	end if
%>
	<span class="textBlockHead">Edit Order Discount</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<form method="post" action="SA_disc_exec.asp" name="form1">
		<tr>
			<td align=right nowrap><b>Discount Code</b></td>
			<td align=left>
				<input type=text name=discCode id=discCode size=20 maxlength=20 value="<%=discCode%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Order Amount From</b></td>
			<td align=left>
				<input type=text name=discFromAmt id=discFromAmt size=10 maxlength=10 value="<%=moneyD(discFromAmt)%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Order Amount To</b></td>
			<td align=left>
				<input type=text name=discToAmt id=discToAmt size=10 maxlength=10 value="<%=moneyD(discToAmt)%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Discount Percentage</b></td>
			<td align=left nowrap>
				<input type=text name=discPerc id=discPerc size=10 maxlength=10 value="<%=discPerc%>"> %
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Discount Amount</b></td>
			<td align=left nowrap>
				<input type=text name=discAmt id=discAmt size=10 maxlength=10 value="<%=moneyD(discAmt)%>">
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Status</b></td>
			<td align=left>
				<select name=discStatus id=discStatus size=1>
					<option value="">-- Select --</option>
					<option value="A" <%=checkMatch(discStatus,"A")%>>Active</option>
					<option value="I" <%=checkMatch(discStatus,"I")%>>InActive</option>
					<option value="U" <%=checkMatch(discStatus,"U")%>>Used</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Use Once Only?</b></td>
			<td align=left>
				<select name=discOnceOnly id=discOnceOnly size=1>
					<option value="">-- Select --</option>
					<option value="Y" <%=checkMatch(discOnceOnly,"Y")%>>Yes</option>
					<option value="N" <%=checkMatch(discOnceOnly,"N")%>>No</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Date Valid From</b></td>
			<td align=left nowrap>
				<%call dateSelect("F",discValidFrom)%>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Date Valid To</b></td>
			<td align=left nowrap>
				<%call dateSelect("T",discValidTo)%>
			</td>
		</tr>
		<tr>
			<td colspan=2 align=center valign=middle>
				<br>
				<input type=hidden name=idDiscOrder id=idDiscOrder value="<%=idDiscOrder%>">
				<input type=hidden name=action      id=action      value="edit">
				<input type=submit name=submit1     id=submit1     value="Update Discount">
				<br><br>
			</td>
		</tr>
		</form>
	</table>
<%
end if

'Add
if action = "add" then
%>
	<span class="textBlockHead">Add Order Discount</span>
	&nbsp;<%call maintNavLinks()%><br><br>
	
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
		<form method="post" action="SA_disc_exec.asp" name="form1">
		<tr>
			<td align=right nowrap><b>Discount Code</b></td>
			<td align=left>
				<input type=text name=discCode id=discCode size=20 maxlength=20>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Order Amount From</b></td>
			<td align=left>
				<input type=text name=discFromAmt id=discFromAmt size=10 maxlength=10>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Order Amount To</b></td>
			<td align=left>
				<input type=text name=discToAmt id=discToAmt size=10 maxlength=10>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Discount Percentage</b></td>
			<td align=left nowrap>
				<input type=text name=discPerc id=discPerc size=10 maxlength=10> %
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Discount Amount</b></td>
			<td align=left nowrap>
				<input type=text name=discAmt id=discAmt size=10 maxlength=10>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Status</b></td>
			<td align=left>
				<select name=discStatus id=discStatus size=1>
					<option value="A">Active</option>
					<option value="I">InActive</option>
					<option value="U">Used</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Use Once Only?</b></td>
			<td align=left>
				<select name=discOnceOnly id=discOnceOnly size=1>
					<option value="Y">Yes</option>
					<option value="N">No</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Date Valid From</b></td>
			<td align=left nowrap>
				<%call dateSelect("F","")%>
			</td>
		</tr>
		<tr>
			<td align=right nowrap><b>Date Valid To</b></td>
			<td align=left nowrap>
				<%call dateSelect("T","")%>
			</td>
		</tr>
		<tr>
			<td colspan=2 align=center valign=middle>
				<br>
				<input type=hidden name=action  id=action  value="add">
				<input type=submit name=submit1 id=submit1 value="Add Discount">
				<br><br>
			</td>
		</tr>
		</form>
	</table>
<%
end if

if action = "edit" or action = "add" then
%>
	<br>
	<span class="textBlockHead">Help and Instructions :</span><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td>

		<b>Discount Code</b> - Mandatory. Unique 
		alpha-numeric code that will be entered by the customer 
		to qualify for a discount. Don't use any special characters 
		or spaces.<br><br>
	
		<b>Order Amount From</b> - Mandatory. The minimum total order 
		amount (excluding taxes and shipping) required to qualify for 
		the discount.<br><br>
	
		<b>Order Amount To</b> - Mandatory. The maximum total order 
		amount (excluding taxes and shipping) required to qualify for 
		the discount.<br><br>
	
		<b>Discount Percentage</b> - Optional if Discount Amount is 
		entered. The percentage of the the total order amount 
		(excluding taxes and shipping) that will be deducted from the 
		order total.<br><br>
		
		<b>Discount Amount</b> - Optional if Discount Percentage is 
		entered. The amount of the the total order amount 
		(excluding taxes and shipping) that will be deducted from 
		the order total.<br><br>
	
		<b>Status</b> - Mandatory. A discount has to be "Active" to be 
		available for use. To prevent a discount from being used, set 
		this to "InActive" or "Used". Discounts that can only be used 
		once (see below), will automatically be set to "Used".<br><br>
	
		<b>Use Once Only?</b> - Mandatory. If set to "Yes", the system 
		will automatically update the status to "Used" after the first 
		time the discount has been applied by any customer.<br><br>
	
		<b>Date Valid From</b> - Mandatory. The date from which the discount 
		will be valid (DD/MM/YYYY).<br><br>
	
		<b>Date Valid To</b> - Mandatory. The date from which the discount 
		will no longer be valid (DD/MM/YYYY).<br><br>
	
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
	<a href=SA_disc.asp?recallCookie=1>List Discounts</a> 
	]
<%
end sub
'*********************************************************************
'Create Date Drop Down boxes
'*********************************************************************
sub dateSelect(FromOrTo,strDate)

	'Declare some variables local to this subroutine
	dim strY, strM, strD

	'Validate From / To indicator
	FromOrTo = UCase(trim(FromOrTo))
	if FromOrTo <> "F" and FromOrTo <> "T" then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid From/To Indicator passed to date routine.")
	end if
	if FromOrTo = "T" then
		FromOrTo = "To"
	else
		FromOrTo = "From"
	end if
	
	'Get date parameter and extract Y, M and D.
	'Default to today's date if invalid or empty.
	if len(strDate) = 8 and isNumeric(strDate) then
		strY = mid(strDate,1,4)
		strM = mid(strDate,5,2)
		strD = mid(strDate,7,2)
	else
		strY = year(now())
		strM = left("00",2-len(datePart("m",now()))) & datePart("m",now())
		strD = left("00",2-len(datePart("d",now()))) & datePart("d",now())
	end if
%>
	<select name="discValid<%=FromOrTo%>DD">
<%
		for I = 1 to 31
			if I < 10 then
%>
				<option value="0<%=I%>" <%=checkMatch(strD,"0" & I)%>>0<%=I%></option>
<%
			else
%>
				<option value="<%=I%>" <%=checkMatch(strD,I)%>><%=I%></option>
<%
			end if
		next
%>
	</select>
	/ 
	<select name="discValid<%=FromOrTo%>MM">
<%
		for I = 1 to 12
			if I < 10 then
%>
				<option value="0<%=I%>" <%=checkMatch(strM,"0" & I)%>>0<%=I%></option>
<%
			else
%>
				<option value="<%=I%>" <%=checkMatch(strM,I)%>><%=I%></option>
<%
			end if
		next
%>
	</select>
	/ 
	<select name="discValid<%=FromOrTo%>YYYY">
<%
		for I = 2002 to 2030
%>
			<option value="<%=I%>" <%=checkMatch(strY,I)%>><%=I%></option>
<%
		next
%>
	</select>
<%
end sub
%>