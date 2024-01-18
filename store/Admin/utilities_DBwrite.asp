<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Check Database Write
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
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<%
'Declare variables
dim mySQL, cn, rs

'*************************************************************************

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Check Database Connection and Write Permissions</font></b>
	<br><br>
</P>

<%
'Open Database
call openDb()

'Create Test Table
mySQL = "CREATE TABLE testWriteTable (testcol INTEGER)"
set rs = openRSexecute(mySQL)

'Delete Test Table
mySQL = "DROP TABLE testWriteTable"
set rs = openRSexecute(mySQL)

'Close Database
call closedb()
%>

<table border=0 cellspacing=0 cellpadding=10 width="100%" class="textBlock">
<tr><td>

	<br>

	<b><font size=3 color=green>Success!</font></b>

	<br><br><br>

	<font size=2>
		The utility was able to connect 
		to your database and make modifications to it.
	</font>

	<br><br>
	
</td></tr>
</table>

<!--#include file="_INCfooter_.asp"-->
