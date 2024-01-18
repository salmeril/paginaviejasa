<%@ Language=VBScript %>
<%
'********************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Administration Home Page
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'********************************************************************
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
'Declare variables
dim mySQL, cn, rs

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
	<b><font size=3>Admin Home</font></b>
	<br><br>
</P>

<table border=0 cellspacing=0 cellpadding=0 width=100%>
	<tr>
		<td align=left valign=top nowrap>
			<b>Current Server Date&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=formatDateTime(date(),vblongDate)%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>Current Server Time&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=time()%>
		</td>
	</tr>
	<tr>
		<td colspan=2 align=center valign=middle>
			<img src="x_pixel.gif" width="100%" height="1" align="absMiddle" vspace=4>
		</td>
	</tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>Orders <%=orderStatusDesc("0")%>&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=numOrders(0,"0") & " ( " & numOrders(-1,"0") & " in the last 24 Hours )"%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>Orders <%=orderStatusDesc("1")%>&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=numOrders(0,"1")%>
		</td>
	</tr>
	<tr>
		<td colspan=2 align=center valign=middle>
			<img src="x_pixel.gif" width="100%" height="1" align="absMiddle" vspace=4>
		</td>
	</tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>HTTP Server&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=Request.ServerVariables("SERVER_SOFTWARE")%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>Default Script Engine&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
<%	
			on error resume next
			Response.Write ScriptEngine _
						 & " - Version " _
						 & ScriptEngineMajorVersion _
						 & "." _
						 & ScriptEngineMinorVersion _
						 & "." _
						 & ScriptEngineBuildVersion
			on error goto 0
%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>MDAC Version&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=cn.Version%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>LCID 1033 Format - Number&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=formatNumber(1234567.89,2)%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>LCID 1033 Format - Date&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=formatDateTime(now(),vbShortDate)%>
		</td>
	</tr>
	<tr>
		<td colspan=2 align=center valign=middle>
			<img src="x_pixel.gif" width="100%" height="1" align="absMiddle" vspace=4>
		</td>
	</tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>Product Image Directory&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=pImagesDir%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>Download Directory&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=pDownloadDir%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>HTTP Scripts Directory&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=urlNonSSL%>
		</td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td align=left valign=top nowrap>
			<b>HTTPS Scripts Directory&nbsp;&nbsp;</b>
		</td>
		<td align=left valign=top>
			<%=urlSSL%>
		</td>
	</tr>

</table>

<br>

<%
'Close database
call closedb()
%>
<!--#include file="_INCfooter_.asp"-->
<%
'********************************************************************
'Get Number of Orders
'********************************************************************
function numOrders(intDays,orderStatus)

	dim tempDate
	
	if intDays = 0 then
		tempDate = dateInt(dateAdd("yyyy",-25,now()))
	else
		tempDate = dateInt(dateAdd("d",intDays,now()))
	end if
	
	mySQL = "SELECT COUNT(*) AS numOrders " _
	      & "FROM   cartHead " _
	      & "WHERE  orderDateInt > '" & tempDate & "' "
	      
	if orderStatus <> "" then
		mySQL = mySQL & "AND orderStatus = '" & orderStatus & "' "
	end if
	
	set rs = openRSexecute(mySQL)
	
	numOrders = rs("numOrders")
	
	call closeRS(rs)
	
end function
%>