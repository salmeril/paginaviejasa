<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Display the General Terms And Conditions for the Shop.
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
<%
'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2

'Session
dim idOrder
dim idCust
'*************************************************************************
%>
<html>
<head>
	<title><%=langGenTOSlink%></title>
	<STYLE type="text/css">
	<!--
	BODY, B, TD, P    
	{
		COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt
	}
	-->
	</STYLE>
</head>

<body>

<p align=center><b><font size=2><%=langGenTOSlink%></font></b></p>

<%
'Open Database Connection
call openDb()

'Get Terms and Conditions
mySQL = "SELECT configValLong " _
	&   "FROM   storeAdmin " _
	&   "WHERE  configVar = 'termsAndCond' " _
	&   "AND    adminType = 'T'"
set rsTemp = openRSexecute(mySQL)
if not rstemp.eof then
	Response.Write trim(rsTemp("configValLong"))
end if
call closeRS(rsTemp)

'Close the Database
call closeDB()		
%>

<p align=center><input type="button" name="SubClose" value="<%=langGenClose%>" onClick="javascript:window.close()"></p>

</body>

</html>