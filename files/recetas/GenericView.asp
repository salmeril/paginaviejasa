<% 
' Generic Database - View Record
' Author: Eli Robillard
' E-Mail: erobillard@ofifc.org
' Revision History:
'  30 Nov  98 - Support for sub-tables
'  24 Nov  98 - Added BGCOLOR=#FFFFFF (white) to the title display, Explorer was overlapping
'				the title with the blue from the rightmost cell
'   1 Oct  98 - Fieldnames beginning with "IMG" will be displayed with the <IMG SRC=""> tag.
'  27 Sept 98 - Added Session("dbType") var. If set to "SQL", the brackets are stripped out of the 
'				SQL queries. 
'  18 Sept 98 - Removed rs.Find references (ADO 2.0) and recreated the functionality with SQL.
' 	9 Sept 98 - Last modified

QUOTE = chr(34)
LT = chr(60)
GT = chr(62)

' Quick security check
If Session("dbDispView") = "" Then 
	Response.Redirect Session("dbViewPage")
End If

' Get info from Session vars (kinda like parameters)
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")

' See if there's a sub-table to display
If Session("dbSubTable") & "x" <> "x" Then
	arrSubTable = Split(Session("dbSubTable"),",")
	IsSubTable = True
End If

' Get the key value of the record to display	
If Request.QueryString("KEY").Count > 0 Then
	strKey = Request.QueryString("KEY")
	Session("dbcurKey") = strKey
Else
	Response.Redirect Session("dbViewPage")
End If

' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn

' Get info from Session vars (kinda like parameters)
strTable = Session("dbRs")
strDisplay = Session("dbDispView")
strKeyField = Session("dbKey")

' Open Recordset and get the field info
strsql = "SELECT * FROM [" & strTable & "]"
If strType = "SQL" Then
	' SQL databases do not allow spaces or brackets in table or field names
	strsql = Replace(strsql,"[","")
	strsql = Replace(strsql,"]","")
End If
set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn
intFieldCount = xrs.Fields.Count
Dim aFields()
ReDim aFields(intFieldCount,4)
For x = 1 to intFieldCount
	aFields(x, 1) = xrs.Fields(x-1).Name 
	aFields(x, 2) = xrs.Fields(x-1).Type 
	aFields(x, 3) = xrs.Fields(x-1).DefinedSize 
Next 

xrs.Close
strsql = "SELECT * FROM [" & strTable & "]"
strsql = strsql & " WHERE [" & aFields(strKeyField,1) & "]" & "=" & strKey
If strType = "SQL" Then
	' SQL databases do not allow spaces or brackets in table or field names
	strsql = Replace(strsql,"[","")
	strsql = Replace(strsql,"]","")
End If
xrs.Open strsql, xConn

If xrs.EOF Then
	Response.Redirect Session("dbViewPage")
End If

' Get the field contents
For x = 1 to intFieldCount
	aFields(x,4) = xrs(x-1)
Next 

xrs.Close
Set xrs = Nothing
xConn.Close
Set xConn = Nothing
%>

<HTML>
<HEAD>
	<TITLE><%=Session("dbTitle")%> - View</TITLE>
</HEAD>
<BODY>
<FONT FACE="Verdana, Arial, Helvetica">

<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC WIDTH="100%">
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<TR>
	<TD BGCOLOR=#99CCCC ALIGN="RIGHT" WIDTH="*"><FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
		<A HREF="<%=Session("dbViewPage")%>">Back to List</A>
	<% If (Session("dbCanEdit") = 1)  and Session("dbKey") > 0 Then %> 
		&nbsp;&nbsp;|&nbsp;&nbsp;	
		<A HREF="<%=Session("dbGenericPath")%>GenericEdit.asp?KEY=<%=aFields(Session("dbKey"),4)%>">Edit</A>
	<% End If %>
	<% If IsSubTable Then %> 
		&nbsp;&nbsp;|&nbsp;&nbsp;	
		<A HREF="<%=Session("dbGenericPath")%>GenericExit.asp?KEY=<%=aFields(Session("dbKey"),4)%>"><%=arrSubTable(0)%></A>
	<% End If %>
	</TD>
</TR>
<TR><TD ALIGN="RIGHT" BGCOLOR=#FFFFFF ROWSPAN=2 WIDTH="1"><FONT SIZE=5 FACE="Verdana, Arial, Helvetica"><STRONG><EM><%=Session("dbTitle")%> - View</EM></STRONG></FONT> </TD></TR>
</TABLE></TD></TR></TABLE>

<P>
<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC>
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<% 	For x = 1 to intFieldCount 
	If Mid(strDisplay, x, 1) = "1" Then 
	%>
		<TR BGCOLOR="#FFFFCC" ALIGN="LEFT">
		<TD >
			<% Response.Write aFields(x,1) %>
		</TD>
		<TD BGCOLOR="White" ALIGN="LEFT">
<%			curVal = aFields(x,4)
			' Blank or null
			If IsNull(curVal) Then
				curVal = "&nbsp;"
			End If
			If Trim(curVal) & "x" = "x" Then
				curVal = "&nbsp;"
			End If
			' Password
			If UCase(Left(aFields(x,1),8)) = "PASSWORD" Then
				curVal = "*****"
			End If
			' Image
			If (UCase(Left(aFields(x,1),3)) = "IMG") Then	
				curVal = LT & "IMG SRC=" & QUOTE & curVal & QUOTE & GT 
			End If
			' Link
			If (UCase(Left(curVal,7)) = "HTTP://") OR (UCase(Left(curVal,7)) = "MAILTO:") Then 
				curVal = LT & "A HREF=" & QUOTE & curVal & QUOTE & GT & curVal & LT & "/A" & GT
			End If
			Response.Write curVal %>
		</TD>		
		</TR>
	<% End If %>
<% Next %>
</TABLE>
</TD></TR>
</TABLE>

<!-- Footer -->
<% If Session("dbFooter") = 1 Then %>
<!--#include file="GenericFooter.inc"-->
<% End If %>
</BODY>
</HTML>
