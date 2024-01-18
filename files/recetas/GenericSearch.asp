<% 
' Generic Database - Search
' Author: Eli Robillard
' E-Mail: erobillard@ofifc.org
' Revision History:
'  30 Nov  98 - File created

QUOTE = chr(34)
LT = chr(60)
GT = chr(62)

' Quick security check, make sure we have an active session
If Session("dbDispList") = "" or Session("dbConn") = "" Then 
	Response.Redirect "GenericError.asp"
End If

' Get info from Session vars (kinda like parameters)
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")

' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn

' Get info from Session vars (kinda like parameters)
strTable = Session("dbRs")
strDisplay = Session("dbDispView")
strKeyField = Session("dbKey")
%>

<HTML>
<HEAD>
	<TITLE><%=Session("dbTitle")%> - Search</TITLE>
</HEAD>
<BODY>
<FONT FACE="Verdana, Arial, Helvetica">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
<TR>
	<TD BGCOLOR=#FFFFFF ROWSPAN=2 WIDTH="1"><FONT SIZE=5 FACE="Verdana, Arial, Helvetica"><STRONG><EM><%=Session("dbTitle")%> - Search</EM></STRONG></FONT> </TD>
	<TD BGCOLOR=#99CCCC ALIGN="RIGHT" WIDTH="*"><FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
		<A HREF="<%=Session("dbViewPage")%>">Back to List</A>
	</TD>
</TR>
<TR><TD ALIGN="RIGHT">&nbsp;</TD></TR>
</TABLE>

<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC>
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<%
' Open Recordset and get the field info
strsql = "SELECT * FROM [" & strTable & "]"
If strType = "SQL" Then
	' SQL databases do not allow spaces or brackets in table or field names
	strsql = Replace(strsql,"[","")
	strsql = Replace(strsql,"]","")
End If
set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn, 1, 2
intFieldCount = xrs.Fields.Count
Dim aFields()
ReDim aFields(intFieldCount,4)
%>
<FORM ACTION="GenericSearchResults.asp" METHOD="POST">
<TR>
	<TH><FONT SIZE=2 FACE="Verdana, Arial, Helvetica">Field</TH>
	<TH><FONT SIZE=2 FACE="Verdana, Arial, Helvetica">Search For</TH>
</TR>
<%
For x = 1 to intFieldCount
	If x mod 2 = 0 Then
		bgcolor="#FFFFCC"
	Else
		bgcolor="White"
	End If
%>
	<TR>
		<TD HEIGHT=0 BGCOLOR=<%= bgcolor %> ALIGN="LEFT">
			<FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
<% 			aFields(x, 1) = xrs.Fields(x-1).Name
			Response.Write aFields(x,1) %>
</TD>
		<TD HEIGHT=0 BGCOLOR=<%= bgcolor %> ALIGN="LEFT">
			<INPUT TYPE="Text" NAME="<%=aFields(x,1)%>" SIZE=40>
		</TD>
	</TR>
<% Next %>
</TABLE>
</TD></TR>
</TABLE>
<INPUT TYPE="Submit" NAME="Submit" VALUE="Search Now">
</FORM>

<P>
<FORM ACTION="GenericSearchResult.asp" METHOD="POST">
Search For: <INPUT TYPE="Text" NAME="strSearch" SIZE=40>
<INPUT TYPE="Submit" NAME="Submit" VALUE="Search Now">
</FORM>
<P>

<%
xrs.Close
Set xrs = Nothing
xConn.Close
Set xConn = Nothing
%>

<!-- Footer -->
<% If Session("dbFooter") = 1 Then %>
<!--#include file="GenericFooter.inc"-->
<% End If %>
</BODY>
</HTML>
