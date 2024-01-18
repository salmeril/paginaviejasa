<% 
' Generic Database - Information Page
' Author: Eli Robillard
' E-Mail: erobillard@ofifc.org
' URL: http://www.ofifc.org/Eli/ASP/GenericArticle.asp
' Revision History:
'  15 Feb  99 - Changed format of title / button bar
'  30 Nov  98 - Added 2 and 4-byte reals, currency to types supported
'  24 Nov  98 - Added BGCOLOR=#FFFFFF (white) to the title display, Explorer was overlapping
'				the title with the blue from the rightmost cell
'   1 Oct  98 - Added a new column to translate the Type # to a name. 
'  27 Sept 98 - Added Session("dbType") var. If set to "SQL", the brackets are stripped out of the 
'				SQL queries. 
' 	9 Sept 98 - Last modified

' Check for an active session
If Session("dbConn") = "" Then 
	Response.Redirect "GenericError.asp"
End If

' Check for user rights
If Session("dbDebug") <> 1 Then 
	Response.Redirect Session("dbViewPage")
End If

' Get info from Session vars 
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
	<TITLE>Generic View - Information</TITLE>
</HEAD>

<BODY>

<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC WIDTH="100%">
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<TR>
	<TD BGCOLOR=#99CCCC ALIGN="RIGHT" WIDTH="*"><FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
		<A HREF="<%=Session("dbGenericPath") & "GenericList.asp"%>">Back to List</A>
	</TD>
</TR>
<TR><TD ALIGN="RIGHT" BGCOLOR=#FFFFFF ROWSPAN=2 WIDTH="1"><FONT SIZE=5 FACE="Verdana, Arial, Helvetica"><STRONG><EM><%=Session("dbTitle")%> - Information</EM></STRONG></FONT> </TD></TR>
</TABLE></TD></TR></TABLE>

<P>
<BR>
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
<TR><TD COLSPAN=5 ALIGN=CENTER BGCOLOR="#FFFFCC"><STRONG>- <%= strTable %> -</STRONG></TD></TR>
<TR>
	<TH><FONT SIZE=2 FACE="Verdana, Arial, Helvetica">#</TH>
	<TH><FONT SIZE=2 FACE="Verdana, Arial, Helvetica">Field</TH>
	<TH><FONT SIZE=2 FACE="Verdana, Arial, Helvetica">Type #</TH>
	<TH><FONT SIZE=2 FACE="Verdana, Arial, Helvetica">Type</TH>
	<TH><FONT SIZE=2 FACE="Verdana, Arial, Helvetica">Length</TH>
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
			<%= x %></TD>
		<% aFields(x, 1) = xrs.Fields(x-1).Name %>
		<TD HEIGHT=0 BGCOLOR=<%= bgcolor %> ALIGN="LEFT">
			<FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
			<%=aFields(x, 1)%></TD>
		<%	aFields(x, 2) = xrs.Fields(x-1).Type %>
		<TD HEIGHT=0 BGCOLOR=<%= bgcolor %> ALIGN="LEFT">
			<FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
			<%=aFields(x, 2)%></TD>
		<%	aFields(x, 3) = xrs.Fields(x-1).DefinedSize %>
		<TD HEIGHT=0 BGCOLOR=<%= bgcolor %> ALIGN="LEFT">
			<FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
<%			Select Case aFields(x,2) 
				Case 2	' 2-Byte Integer %>
					2-Byte Integer
				<% Case 3 ' 4-Byte Integer %>
					4-Byte Integer
				<% Case 4 ' 2-Byte Real %>
					2-Byte Real
				<% Case 5 ' 4-Byte Real %>
					4-Byte Real
				<% Case 6 ' Currency %>
					Currency
				<% Case 11 	' Boolean True/False %>
					True/False
				<% Case 135 ' Date / Time Stamp, usually created with the Now() function %>
					Date / Time
				<% Case 200 ' String %>
					String
				<% Case 201 ' Memo %>
					Memo
<% 			End Select %>
		</TD>
		<TD BGCOLOR=<%= bgcolor %> ALIGN="LEFT">
			<FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
			<%=aFields(x, 3)%></TD>
	</TR>
<% Next %>
</TABLE>
</TD></TR>
</TABLE>
<%
xrs.Close
Set xrs = Nothing
xConn.Close
Set xConn = Nothing
%>
<P>
Connection: <STRONG><%= strConn %></STRONG><BR>
Where: <STRONG><%= Session("dbWhere") %></STRONG>

<!-- Footer -->
<% If Session("dbFooter") = 1 Then %>
<!--#include file="GenericFooter.inc"-->
<% End If %>
</BODY>
</HTML>
