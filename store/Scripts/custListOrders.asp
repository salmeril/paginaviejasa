<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : List Customer Orders
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
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="_INCappFunctions_.asp"-->
<%
'cartHead
dim orderStatus
dim orderDate
dim Total
dim shipmentMethod
dim paymentType

'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2

'Session
dim idOrder
dim idCust

'*************************************************************************

'Open Database Connection
call openDb()

'Store Configuration
if loadConfig() = false then
	call errorDB(langErrConfig,"")
end if

'Get/Set Cart/Order Session
idOrder = sessionCart()

'Get/Set Customer Session
idCust = sessionCust()

'Double-check that the Customer is now "logged in"
if isNull(idCust) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrNotLoggedIn)
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<!-- Outer Table Cell -->
<table border="0" cellpadding="0" cellspacing="0" width="450"><tr><td>

<!-- Heading -->
<table border="0" cellspacing="0" cellpadding="2" width="100%">
	<tr><td valign=middle class="CPpageHead">
		<b><%=langGenYourAccount%></b><br>
	</td></tr>
</table>

<!-- Main table -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr>
		<td colspan=4>&nbsp;</td>
	</tr>
	<tr>
		<td colspan=4 align=left valign=middle nowrap class="CPgenHeadings">
			&raquo;&nbsp;<b><a href="20_Customer.asp?action=modify"><%=langGenModifyAccInfo%></a></b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&raquo;&nbsp;<b><a href="10_Logon.asp?action=logoff"><%=langGenLogoff%></a></b>
		</td>
	</tr>
	<tr>
		<td colspan=4>&nbsp;</td>
	</tr>
	<tr>
		<td colspan=4 align=left>
			<%=langGenClickOnOrder%>
		</td>
	</tr>
	<tr> 
		<td class="CPgenHeadings" nowrap align=left><b><%=langGenOrderNumber%></b></td>
		<td class="CPgenHeadings" nowrap align=left><b><%=langGenOrderDate%></b></td>
		<td class="CPgenHeadings" nowrap align=left><b><%=langGenTotal%></b></td>
		<td class="CPgenHeadings" nowrap align=left><b><%=langGenOrderStatus%></b></td>
	</tr>
<%
	'Retrieve Customer's Orders
	mySQL = "SELECT idOrder,orderStatus,orderDate,randomKey,Total " _
		  & "FROM   cartHead " _
		  & "WHERE  idCust = " & validSQL(idCust,"I") & " " _
		  & "AND    orderStatus <> 'U' " _
		  & "ORDER BY orderDate DESC "
	set rsTemp = openRSexecute(mySQL)
	if not rstemp.eof then
		do while not rsTemp.EOF
%>
			<tr>
<%
				if UCase(rsTemp("orderStatus")) = "S" then
%>
				<td nowrap align=left valign=top width="25%">
					(<%=pOrderPrefix & "-" & rsTemp("idOrder")%>)
				</td>
				<td nowrap align=left valign=top width="25%">
					<%=formatTheDate(rsTemp("orderDate"))%>
				</td>
				<td nowrap align=left valign=top width="25%">
					<%=langGenNotApplicable%>
				</td>
<%
				else
%>
				<td nowrap align=left valign=top width="25%">
					<a href="custViewOrders.asp?idOrder=<%=rsTemp("idOrder")%>"><%=pOrderPrefix & "-" & rsTemp("idOrder")%></a>
				</td>
				<td nowrap align=left valign=top width="25%">
					<%=formatTheDate(rsTemp("orderDate"))%>
				</td>
				<td nowrap align=left valign=top width="25%">
					<%=pCurrencySign & moneyS(rsTemp("Total"))%>
				</td>
<%
				end if
%>
				<td nowrap align=left valign=top width="25%">
<%
					Response.Write orderStatusDesc(rsTemp("orderStatus"))
					if UCase(rsTemp("orderStatus")) = "S" then
%>
						&nbsp;( <a href="<%=urlNonSSL%>05_Gateway.asp?action=retrieve&idOrder=<%=rsTemp("idOrder")%>&randomKey=<%=rsTemp("randomKey")%>"><%=langGenRetrieveCart%></a> )
<%
					end if
%>
				</td>
			</tr>
<%
			rsTemp.MoveNext
		loop
	else
%>
		<tr>
			<td colspan=4 align=left>
				<br>
				<b><font color=red><%=langErrNoOrders%></font></b>
				<br><br>
			</td>
		</tr>
<%
	end if
	call closeRS(rsTemp)
%>

</table>

<!-- End Outer Table Cell -->
</td></tr></table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%
call closedb()
%>
