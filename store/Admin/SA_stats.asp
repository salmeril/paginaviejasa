<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Statistics & Reports
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
'Work fields
dim repType
dim repPeriod
dim repInterval
dim repOrderStatus
dim startDate
dim startDateAll
dim endDate
dim aValues
dim aLabels
dim strTitle
dim strYAxisLabel
dim strXAxisLabel
dim I
dim repID
dim errMsg

'Database
dim mySQL, cn, rs, rs2

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Get form values
repID = trim(Request.Form("repID"))
if repID = "01" then		'ORDERS
	repType			= lCase(trim(Request.Form("repType")))
	repOrderStatus	= lCase(trim(Request.Form("repOrderStatus")))
	repPeriod		= lCase(trim(Request.Form("repPeriod")))
	repInterval		= lCase(trim(Request.Form("repInterval")))
elseif repID = "02" then	'GENERAL REPORTS
	repType			= lCase(trim(Request.Form("repType")))
end if
%>
<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Statistics</font></b>
	<br><br>
</P>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">

	<tr>
		<form method="post" action="SA_stats.asp" name="form2">
		<td align=right valign=top nowrap>
			<select name=repType id=repType size=1>
				<option value="PQ" <%=checkMatch(repType,"PQ")%>>Top Products by Quantity Sold</option>
				<option value="PA" <%=checkMatch(repType,"PA")%>>Top Products by Amount Sold</option>
				<option value="CQ" <%=checkMatch(repType,"CQ")%>>Top Customers by Number of Orders</option>
				<option value="CA" <%=checkMatch(repType,"CA")%>>Top Customers by Order Value</option>
				<option value="LA" <%=checkMatch(repType,"LA")%>>Top Countries by Order Value</option>
				<option value="IL" <%=checkMatch(repType,"IL")%>>Products with the lowest Inventory levels</option>
			</select>&nbsp;
			<input type=hidden name=repID   id=repID   value="02">
			<input type=submit name=submit1 id=submit1 value="Show">
		</td>
		</form>
	</tr>

	<tr>
		<form method="post" action="SA_stats.asp" name="form1">
		<td align=right valign=top nowrap>
			<select name=repType id=repType size=1>
				<option value="N" <%=checkMatch(repType,"N")%>>Number of</option>
				<option value="T" <%=checkMatch(repType,"T")%>>Total amount for</option>
			</select>&nbsp;
			<select name=repOrderStatus id=repOrderStatus size=1>
				<option value="U" <%=checkMatch(repOrderStatus,"U")%>>Unfinalized</option>
				<option value="S" <%=checkMatch(repOrderStatus,"S")%>>Saved</option>
				<option value="0" <%=checkMatch(repOrderStatus,"0")%>>Pending</option>
				<option value="1" <%=checkMatch(repOrderStatus,"1")%>>Paid</option>
				<option value="2" <%=checkMatch(repOrderStatus,"2")%>>Shipped</option>
				<option value="7" <%=checkMatch(repOrderStatus,"7")%>>Complete</option>
				<option value="9" <%=checkMatch(repOrderStatus,"9")%>>Cancelled</option>
			</select>&nbsp;
			orders for the last 
			<select name=repPeriod id=repPeriod size=1>
				<option value="1"  <%=checkMatch(repPeriod,"1") %>>1</option>
				<option value="2"  <%=checkMatch(repPeriod,"2") %>>2</option>
				<option value="3"  <%=checkMatch(repPeriod,"3") %>>3</option>
				<option value="4"  <%=checkMatch(repPeriod,"4") %>>4</option>
				<option value="5"  <%=checkMatch(repPeriod,"5") %>>5</option>
				<option value="6"  <%=checkMatch(repPeriod,"6") %>>6</option>
				<option value="7"  <%=checkMatch(repPeriod,"7") %>>7</option>
				<option value="8"  <%=checkMatch(repPeriod,"8") %>>8</option>
				<option value="9"  <%=checkMatch(repPeriod,"9") %>>9</option>
				<option value="10" <%=checkMatch(repPeriod,"10")%>>10</option>
				<option value="11" <%=checkMatch(repPeriod,"11")%>>11</option>
				<option value="12" <%=checkMatch(repPeriod,"12")%>>12</option>
				<option value="13" <%=checkMatch(repPeriod,"13")%>>13</option>
				<option value="14" <%=checkMatch(repPeriod,"14")%>>14</option>
			</select>&nbsp;
			<select name=repInterval id=repInterval size=1>
				<option value="D"  <%=checkMatch(repInterval,"D") %>>Day(s)</option>
				<option value="WW" <%=checkMatch(repInterval,"WW")%>>Week(s)</option>
				<option value="M"  <%=checkMatch(repInterval,"M") %>>Month(s)</option>
			</select>&nbsp;
			<input type=hidden name=repID   id=repID   value="01">
			<input type=submit name=submit1 id=submit1 value="Show">
		</td>
		</form>
	</tr>
	
</table>
<%
'Display the requested chart
if repID = "01" then		'ORDERS
	call orderRep()
	call ShowChart(aValues,aLabels,strTitle,strXAxisLabel,strYAxisLabel)
elseif repID = "02" then	'GENERAL REPORTS
	call generalRep()
	call ShowTable(aValues,aLabels,strTitle)
else
%>
	<br><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr>
		<td align=center>
			<br>To create a report, select the query parameters you want and click the "Show" button.<br><br>
		</td>
	</tr>
	</table>
<%
end if

'If there was an error, display it
if len(trim(errMsg)) > 0 then
%>
	<br><br>
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr><td align=center><br><span class="errMsg"><%=errMsg%></span><br><br></td></tr>
	</table>
<%
end if

%>
<!--#include file="_INCfooter_.asp"-->
<%

'Close Database Connection
call closedb()

'*********************************************************************
'ORDER REPORTS
'*********************************************************************
Sub orderRep()
	
	'Determine start date of FIRST interval
	startDate = dateAdd(repInterval,CLng("-"&repPeriod),now())
	if repInterval = "m" then							'Adjust Day to 1st
		startDate = datePart("m",startDate) & "/01/" & datePart("yyyy",startDate)
	end if
	startDateAll = startDate

	'Determine end date of FIRST interval
	if repInterval = "m" or repInterval = "ww" then		'Months or Weeks
		endDate = dateAdd(repInterval,1,startDate)		'Add Interval
		endDate = dateAdd("d",-1,endDate)				'Subtract 1 day
	else												'Days
		endDate = startDate								'EndDate = StartDate
	end if
	
	'Get the required data from the database
	aValues = ""
	aLabels = ""
	for I = 0 to CLng(repPeriod) 'Starting from 0 includes current day/week/month

		'Get values from Database
		if repType = "n" then
			mySQL = "SELECT COUNT(*)   as orderTotal "
		else
			mySQL = "SELECT SUM(total) as orderTotal "
		end if
		mySQL = mySQL _
		      & "FROM   cartHead " _
		      & "WHERE  orderStatus='"   & repOrderStatus & "' " _
		      & "AND    orderDateInt>='" & left(dateInt(startDate),8) & "000000' " _
		      & "AND    orderDateInt<='" & left(dateInt(endDate),8) & "999999' "
		set rs = openRSexecute(mySQL)
		if not isNumeric(rs("orderTotal")) then
			aValues = aValues & "0"
		else
			aValues = aValues & rs("orderTotal")
		end if
		call closeRS(rs)

		'Determine what label to show
		if repInterval = "m" or repInterval = "ww" then	'Months or Weeks
			aLabels = aLabels & formatTheDate(startDate) & "-" & formatTheDate(endDate)
		else											'Days
			aLabels = aLabels & formatTheDate(startDate)
		end if
		
		'Tack on a comma to the end
		if I <> CLng(repPeriod) then
			aValues = aValues & ","
			aLabels = aLabels & ","
		end if
		
		'Increment Start and End dates
		startDate = dateAdd("d",1,endDate)
		if repInterval = "m" or repInterval = "ww" then	'Months or Weeks
			endDate = dateAdd(repInterval,1,startDate)	'Add Interval
			endDate = dateAdd("d",-1,endDate)			'Subtract 1 day
		else											'Days
			endDate = startDate							'EndDate = StartDate
		end if
		
	next

	'Make Arrays
	aValues = split(aValues,",")
	aLabels = split(aLabels,",")
	
	'Report Title
	strTitle = "<b>From</b> <i>" & formatTheDate(startDateAll) & "</i> <b>to</b> <i>" & formatTheDate(now()) & "</i>"

	'Interval Label (Y Axis)
	select case repInterval
	case "d"
		strYAxisLabel = "Days"
	case "ww"
		strYAxisLabel = "Weeks"
	case "m"
		strYAxisLabel = "Months"
	end select

	'Report Type Label (X Axis)
	select case repType
	case "n"
		strXAxisLabel = "Number of Orders"
	case "t"
		strXAxisLabel = "Order Total Amounts"
	end select
	
end sub

'*********************************************************************
'GENERAL REPORTS
'*********************************************************************
Sub generalRep()
	
	'Get the required data from the database
	select case repType
		case "pq"
			mySQL = "SELECT a.description, SUM(a.quantity) AS prodQty " _
			      & "FROM   cartRows a, cartHead b " _
			      & "WHERE  a.idOrder = b.idOrder " _
			      & "AND   (b.orderStatus='1' OR b.orderStatus='2' OR b.orderStatus='7') " _
			      & "GROUP BY a.description " _
			      & "ORDER BY SUM(a.quantity) DESC "
			strTitle = "Top Products by Quantity Sold"
			aLabels  = array("Description","Quantity")
		case "pa"
			mySQL = "SELECT a.description, SUM(a.quantity*a.unitPrice) AS prodTotal " _
			      & "FROM   cartRows a, cartHead b " _
			      & "WHERE  a.idOrder = b.idOrder " _
			      & "AND   (b.orderStatus='1' OR b.orderStatus='2' OR b.orderStatus='7') " _
			      & "GROUP BY a.description " _
			      & "ORDER BY SUM(a.quantity*a.unitPrice) DESC "
			strTitle = "Top Products by Amount Sold"
			aLabels  = array("Description","Amount")
		case "cq"
			mySQL = "SELECT lastName, name, COUNT(*) AS custQty " _
			      & "FROM   cartHead " _
			      & "WHERE  orderStatus='1' OR orderStatus='2' OR orderStatus='7' " _
			      & "GROUP BY lastName,name " _
			      & "ORDER BY COUNT(*) DESC "
			strTitle = "Top Customers by Number of Orders"
			aLabels  = array("Last Name","Name","Orders")
		case "ca"
			mySQL = "SELECT lastName, name, SUM(total) AS custTotal " _
			      & "FROM   cartHead " _
			      & "WHERE  orderStatus='1' OR orderStatus='2' OR orderStatus='7' " _
			      & "GROUP BY lastName,name " _
			      & "ORDER BY SUM(total) DESC "
			strTitle = "Top Customers by Order Value"
			aLabels  = array("Last Name","Name","Total")
		case "la"
			mySQL = "SELECT locCountry, SUM(total) AS cntryTotal " _
			      & "FROM   cartHead " _
			      & "WHERE  orderStatus='1' OR orderStatus='2' OR orderStatus='7' " _
			      & "GROUP BY locCountry " _
			      & "ORDER BY SUM(total) DESC "
			strTitle = "Top Countries by Order Value"
			aLabels  = array("Country","Total")
		case "il"
			mySQL = "SELECT description, stock " _
			      & "FROM   products " _
			      & "ORDER BY stock ASC "
			strTitle = "Products with the lowest Inventory levels"
			aLabels  = array("Description","Stock")
	end select
	set rs = openRSexecute(mySQL)
	if not rs.eof then
		aValues = rs.getRows(50)
	end if
	call closeRS(rs)

end sub

'*********************************************************************
'Create a graph of passed values and labels
'*********************************************************************
Sub ShowChart(ByRef aValues, ByRef aLabels, ByRef strTitle, ByRef strXAxisLabel, ByRef strYAxisLabel)

	'Constants
	Const barHeight   = 15
	Const maxBarWidth = 400
	Const colMain     = "#FFFFFF"
	Const colXYLabels = "#DDDDDD"
	Const colLabels   = "#EEEEEE"
	'Variables
	dim maxValue, I, tot1, tot2
	
	'Do some checks on the parms
	if not (IsArray(aValues) and IsArray(aLabels)) then
		errMsg = "Invalid Value or Label array."
		exit sub
	end if
	if UBound(aValues) <> UBound(aLabels) then
		errMsg = "Value and Label array not the same length."
		exit sub
	end if

	'Determine Max Value
	maxValue = 0
	For I = 0 To UBound(aValues)
		If Cdbl(maxValue) < CDbl(aValues(I)) Then 
			maxValue = aValues(I)
		end if
	Next
	
	'If Max value = 0, then there is nothing to display
	if CDbl(maxValue) = 0 then
		errMsg = "There is no data for the current selection."
		exit sub
	end if
%>
	<br>
	<table border=0 cellspacing=0 cellpadding=2 bgcolor="<%=colMain%>">
		<tr>
			<td rowspan=<%=UBound(aValues)+3%> width=0 align=center valign=middle bgcolor="<%=colXYLabels%>">
<%
			for I = 1 to len(strYAxisLabel)
				Response.Write "<b>" & mid(strYAxisLabel,I,1) & "</b><br>"
			next
%>
			</td>
			<td colspan=2 align=center valign=middle bgcolor="<%=colXYLabels%>"><%=strTitle%></td>
		</tr>
<%
		for I = 0 to Ubound(aValues)
			'1) Calculate the value as a percentage of the maximum 
			'   value.
			'2) Determine the number of pixels the percentage 
			'   represents as a portion of the maximum allowed bar 
			'   width.
			tot1 = Int((CDbl(aValues(I)) / Cdbl(maxValue)) * 100)
			tot2 = Int((tot1 * maxBarWidth) / 100)
			
			'Should we show the decimals?
			if repType = "t" then
				aValues(I) = moneyD(aValues(I))
			end if
			
			'If tot2 is 0 change it to 1, otherwise the width tag for
			'the image is not displayed properly in NS4.7
			if tot2=0 then 
				tot2=1 
			end if
%>
			<tr>
				<td align=right valign=middle nowrap bgcolor="<%=colLabels%>"><%=replace(aLabels(I)," ","&nbsp;")%></td>
				<td align=left  valign=middle nowrap><img src="x_red.gif" border=0 width=<%=tot2%> height=<%=barHeight%> align=absMiddle>&nbsp;<i><%=aValues(I)%></i></td>
			</tr>
<%
		next
%>
		<tr>
			<td bgcolor="<%=colXYLabels%>">&nbsp;</td>
			<td align=center valign=middle nowrap bgcolor="<%=colXYLabels%>"><b><%=strXAxisLabel%></b></td>
		</tr>
	</table>
<%	
End Sub
'*********************************************************************
'Create a table of passed values and labels
'*********************************************************************
Sub ShowTable(ByRef aValues, ByRef aLabels, ByRef strTitle)

	'Variables
	dim rowColor, col1, col2
	dim row, col

	'Do some checks on the parms
	if not IsArray(aValues) then
		errMsg = "There is no data for the current selection."
		exit sub
	end if
%>
	<p align=center><span class="textBlockHead"><%=strTitle%></span></p>
	
	<table border=0 cellspacing=0 cellpadding=5 width="100%" class="listTable">
		<tr>
<%
			for col = 0 to UBound(aLabels)
				Response.Write "<td class=listRowHead><b>" & aLabels(col) & "</b></td>"
			next
%>
		</tr>
<%
		'Row Colors
		col1 = "#DDDDDD"
		col2 = "#EEEEEE"
		rowColor = col2
		
		'Write Rows
		for row = 0 to UBound(aValues,2)

			'Write columns
			Response.Write "<tr>"
			for col = 0 to UBound(aValues,1)
				Response.Write "<td valign=top bgcolor=" & rowColor & ">" & aValues(col,row) & "</td>"
			next
			Response.Write "</tr>"

			'Switch Row Color
			if rowColor = col2 then
				rowColor = col1
			else
				rowColor = col2
			end if
			
		next
%>
	</table>
<%	
End Sub
%>