<% 
' Generic Database - Delete Record
' Author: Eli Robillard
' E-Mail: erobillard@ofifc.org
' URL: http://www.ofifc.org/Eli/ASP/GenericArticle.asp
' Revision History:
'  15 Feb  99 - When opening recordset, using Dynamic Recordset (2) and Optimistic locking (3),
'			  - Support for Confirm Delete. Use Session("dbConfirmDelete") = 1 to activate.
'				Implementation based on suggestions of Helge Larsen.
'  30 Sept 98 - After rs.Delete, rs.Update was not required before closing.
'  27 Sept 98 - Added Session("dbType") var. If set to "SQL", the brackets are stripped out of the 
'				SQL queries. 
'  18 Sept 98 - Removed rs.Find references (ADO 2.0) and recreated the functionality with SQL.
' 	9 Sept 98 - Last modified

' Check for an active session
If Session("dbConn") = "" Then 
	Response.Redirect "GenericError.asp"
End If

' Check user rights
If Session("dbCanDelete") <> 1 Then 
	Response.Redirect Session("dbViewPage")
End If

' Get info from Session vars (kinda like parameters)
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")
strDisplay = Session("dbDispView")
strKeyField = Session("dbKey")

' Default to no confirmation
DoConfirm = False

' If we don't get passed a key to delete or there's no unique key field defined then get out.
If (Request.QueryString("KEY").Count = 0) OR (strKeyField = "") Then
	Response.Redirect Session("dbViewPage")
Else
	strKey = Request.QueryString("KEY")
	Session("dbcurKey") = strKey
End If

If Request.QueryString("CMD").Count > 0 Then
	' See if we need to confirm the deletion
	strCMD = Request.QueryString("CMD")
	If strCMD = "'CON'" Then
		DoConfirm = True
	End If
End If

' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn

' Open Recordset and get the field info
set xrs =  Server.CreateObject("ADODB.Recordset")
xrs.Open strTable, xConn, 1, 2
intFieldCount = xrs.Fields.Count
Dim aFields()
ReDim aFields(intFieldCount,4)
For x = 1 to intFieldCount
	aFields(x, 1) = xrs.Fields(x-1).Name 
	aFields(x, 2) = xrs.Fields(x-1).Type 
	aFields(x, 3) = xrs.Fields(x-1).DefinedSize 
Next 

xrs.Close
strsql = "SELECT * FROM [" & strTable & "]" & " WHERE [" & aFields(strKeyField,1) & "]" & "=" & strKey
If strType = "SQL" Then
	' SQL databases do not allow spaces or brackets in table or field names
	strsql = Replace(strsql,"[","")
	strsql = Replace(strsql,"]","")
End If
xrs.Open strsql, xConn, 2, 3

' Get the field contents
For x = 1 to intFieldCount
	aFields(x,4) = xrs(x-1)
Next 

If DoConfirm Then %>
<HTML>
<HEAD>
	<TITLE><%=Session("dbTitle")%> - Delete?</TITLE>
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
<TR><TD ALIGN="RIGHT" BGCOLOR=#FFFFFF ROWSPAN=2 WIDTH="1"><FONT SIZE=5 FACE="Verdana, Arial, Helvetica"><STRONG><EM><%=Session("dbTitle")%> - Delete</EM></STRONG></FONT> </TD></TR>
</TABLE></TD></TR></TABLE>

<P>
<CENTER>
<FONT SIZE=5 FACE="Verdana, Arial, Helvetica"><STRONG><EM>Delete this record?</EM></STRONG></FONT><BR>
<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC>
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<TR>
<TD BGCOLOR="#FFFFCC"><FONT SIZE=3 FACE="Verdana, Arial, Helvetica"><A HREF="<%= Session("dbGenericPath") %>GenericDelete.asp?KEY=<%= strKey %>">YES</A></FONT></TD>
<TD BGCOLOR="#FFFFCC"><FONT SIZE=3 FACE="Verdana, Arial, Helvetica"><A HREF="<%=Session("dbViewPage")%>"><STRONG>NO</STRONG></A></FONT></TD>
</TR>
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
			If UCase(Left(curVal,7)) = "HTTP://" Then 	
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
</CENTER>

</BODY>
</HTML>

<%
	Set xrs = Nothing
	xConn.Close
	Set xConn = Nothing
Else

	If xrs.EOF Then
		Response.Redirect Session("dbViewPage")
	End If

	xrs.Delete
	xrs.Close
	Set xrs = Nothing
	xConn.Close
	Set xConn = Nothing

	Response.Redirect Session("dbViewPage")
End If
%>

