<% 
' Generic Database - List Records
' Author: Eli Robillard
' E-Mail: erobillard@ofifc.org
' URL: http://www.ofifc.org/Eli/ASP/GenericArticle.asp
' Revision History:
'  19 Feb  99 - Don't show a link if dbURLfor field is blank. 
'  15 Feb  99 - Changed format of title / button bar
'			  - Support for Confirm Delete. Use Session("dbConfirmDelete") = 1 to activate.
'				Implementation based on suggestions of Helge Larsen.
'  30 Nov  98 - Support for sub-tables
'  24 Nov  98 - Added BGCOLOR=#FFFFFF (white) to the title display, Explorer was overlapping
'				the title with the blue from the rightmost cell
'			  - Added the dbURLforX var. Used in the Config File, this allows one field to 
'               be wrapped in the URL stored in another field. See article for syntax.
'   9 Oct  98 - Implemented Prev/Next buttons and a setting for # of Recs to display per Page 
'   1 Oct  98 - Fieldnames beginning with "IMG" will be displayed with the <IMG SRC=""> tag.
'  27 Sept 98 - Added Session("dbType") var. If set to "SQL", the brackets are stripped out of the 
'				SQL queries. 
'  11 Sept 98 - Database drivers don't allow sorting on a memo field. Added a check
' 				to prevent the problem while displaying the headers. 
'  11 Sept 98 - Fixed column headers, changing the REF to Session("dbViewPage") broke 
'				the sorting feature if the dbViewPage was set to the Config File.
' 	9 Sept 98 - Last modified

QUOTE = chr(34)
LT = chr(60)
GT = chr(62)

' Quick security check, make sure we have an active session
If Session("dbDispList") = "" or Session("dbConn") = "" Then 
	Response.Redirect "GenericError.asp"
End If

' Get the parameters set in the Config File
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")
strDisplay = Session("dbDispList")
strSearchFields = Session("dbSearchFields")
strWhere = Session("dbWhere")
intPrimary = Session("dbKey")
intOrderBy = Session("dbOrder")

' See if there's a sub-table to display
If Session("dbSubTable") & "x" <> "x" Then
	arrSubTable = Split(Session("dbSubTable"),",")
	IsSubTable = True
End If
' Check to see if the Order was specified by a parameter
If Request.QueryString("ORDER").Count > 0 Then
	intOrderBy = Request.QueryString("ORDER")
	Session("dbOrder") = intOrderBy
End If
' Check for a limit on Records per Page
If Session("dbRecsPerPage") > 0 Then
	intDisplayRecs = Session("dbRecsPerPage")
Else
	intDisplayRecs = 10000
End If
' Check for a START parameter
If Request.QueryString("START").Count > 0 Then
	intStartRec = Request.QueryString("START")
	Session("dbStartRec") = intStartRec
Else
	' Check for a StartRec variable in the Config File
	If Session("dbStartRec") > 0 Then
		intStartRec = Session("dbStartRec")
	Else
		intStartRec = 1
	End If
End If
'Set the last record to display
intStopRec = intStartRec + intDisplayRecs - 1

' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn

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
ReDim aFields(intFieldCount,3)
For x = 1 to intFieldCount
	aFields(x, 1) = xrs.Fields(x-1).Name 
	aFields(x, 2) = xrs.Fields(x-1).Type 
	aFields(x, 3) = xrs.Fields(x-1).DefinedSize 
Next 

xrs.Close
Set xrs = Nothing
%>

<HTML>
<HEAD>
	<TITLE><%= Session("dbTitle") %></TITLE>
</HEAD>

<BODY>
<FONT FACE="Verdana, Arial, Helvetica">
<%
' Reopen the Recordset using the parameters passed
strsql = "SELECT * FROM [" & strTable & "]"
If NOT (strWhere = "") Then
	strsql = strsql & " WHERE " & strWhere
End If
If intOrderBy > 0 Then 
	strsql = strsql & " ORDER BY [" & aFields(intOrderBy, 1) & "]"
End If
If strType = "SQL" Then
	' SQL databases do not allow spaces or brackets in table or field names
	strsql = Replace(strsql,"[","")
	strsql = Replace(strsql,"]","")
End If
set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn
' xrs.MoveFirst
%>

<!-- Main Body Start -->
<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC WIDTH="100%">
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<TR>
	<TD BGCOLOR=#99CCCC ALIGN="RIGHT" WIDTH="*"><FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
		<A HREF="<%=Session("dbGenericPath")%>GenericExit.asp">Back</A>
		<% If Session("dbCanAdd") = 1 Then %>
			&nbsp;&nbsp;|&nbsp;&nbsp;
			<A HREF="<%=Session("dbGenericPath")%>GenericEdit.asp?CMD=NEW">Add&nbsp;New</A>
		<% End If %>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<A HREF="<%=Session("dbViewPage")%>">Reset</A>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<A HREF="<%=Session("dbGenericPath")%>GenericList.asp">Refresh</A>
		<% If Session("dbDebug") = 1 Then %>
			&nbsp;&nbsp;|&nbsp;&nbsp;
			<A HREF="<%=Session("dbGenericPath")%>GenericInfo.asp">db&nbsp;Info</A>
		<% End If %>
	</TD>
</TR>
<TR><TD ALIGN="RIGHT" BGCOLOR=#FFFFFF ROWSPAN=2 WIDTH="1"><FONT SIZE=5 FACE="Verdana, Arial, Helvetica"><STRONG><EM><%=Session("dbTitle")%></EM></STRONG></FONT> </TD></TR>
</TABLE></TD></TR></TABLE>

<P>
<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC>
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<TR>
<% For x = 1 to intFieldCount
		' If the field is to be displayed then
		If Mid(strDisplay, x, 1) = "1" Then 
			strConn = "ORDER=" & x 
			If aFields(x,2) = 201 Then
			' If the field type is a BLOB, then don't display it as a sortable field
%>
				<TH><%=aFields(x, 1)%></TH>
			<% Else %>
				<TH><A HREF="GenericList.asp?<%=strConn%>"><%= aFields(x, 1) %></A></TH>
			<% End If
		End If
	Next
%>
<!-- 	<TH>&nbsp;</TH> -->
</TR>
<%
intCount = 0
intActual = 0
Do While (NOT xrs.EOF) AND (intCount < intStopRec)
	intCount = intCount + 1
	If Cint(intCount) >= Cint(intStartRec) Then 
		intActual = intActual + 1
%>
<TR>
	<%	x = 0
		For Each xField in xrs.Fields
			x = x + 1
			curVal = xField.Value
			' Every other line will have a shaded background
			If intCount mod 2 = 0 Then
				bgcolor="#FFFFCC"
			Else
				bgcolor="White"
			End If
			' If on the Key field, build the link used to load the Viewer, Editor, or Deleter
			If x = CInt(intPrimary) Then
				Session("zcurTable") = strTable
				Session("zcurDisplay") = strDisplay
				Session("zcurKeyField") = aFields(x,1)
				strLink = "KEY=" & xField.Value
			End If
			If Mid(strDisplay, x, 1) = "1" Then %>
<TD BGCOLOR=<%= bgcolor %> ALIGN="LEFT">
<%				If IsNull(curVal) Then
					' Empty or Null
					curVal = "&nbsp;"
				End If
				If UCase(Left(aFields(x,1),8)) = "PASSWORD" Then
					' password field
					curVal = "*****"
				End If
				If (UCase(Left(aFields(x,1),3)) = "IMG") Then
					' image field
					If Session("dbMaxImageSize") = 0 Then
						curVal = LT & "IMG SRC=" & QUOTE & curVal & QUOTE & GT 
					Else 
						curVal = LT & "IMG SRC=" & QUOTE & curVal & QUOTE & " WIDTH=" & QUOTE & Session("dbMaxImageSize") & QUOTE & GT 
					End If
				End If		
				If (UCase(Left(curVal,7)) = "HTTP://") OR (UCase(Left(curVal,7)) = "MAILTO:") Then 
					' contents is an URL
					curVal = LT & "A HREF=" & QUOTE & curVal & QUOTE & GT & xField.Value & LT & "/A" & GT
 				End If
				
				strContainsURL = "dbURLfor" & CStr(x)
				If Session(strContainsURL) > 0 Then
					If Not (Trim(xrs(aFields(Session(strContainsURL),1))) & "x" = "x") Then
						strURL = xrs(aFields(Session(strContainsURL),1))
						curVal = "<A HREF=" & QUOTE & strURL & QUOTE & ">" & curVal & "</A>"
					End If
				End If
%>
<%= curVal %>&nbsp;</TD>
<% 			End If 
		Next 
%>

<%		If IsSubTable Then %>
			<TD BGCOLOR=<%= bgcolor %> ALIGN="CENTER"><A HREF="<%=Session("dbGenericPath")%>GenericExit.asp?<%=strLink%>"><%=arrSubTable(0)%></A></TD>
<%		End if
		If (Session("dbDispView") <> "")  and Session("dbKey") > 0 Then %>
			<TD BGCOLOR=<%= bgcolor %> ALIGN="CENTER"><A HREF="<%= Session("dbGenericPath") %>GenericView.asp?<%=strLink%>">View</A></TD>
<% 		End If 
		If (Session("dbCanEdit") = 1)  and Session("dbKey") > 0 Then %>
			<TD BGCOLOR=<%= bgcolor %> ALIGN="CENTER"><A HREF="<%= Session("dbGenericPath") %>GenericEdit.asp?<%=strLink%>">Edit</A></TD>
<% 		End If 
		If (Session("dbCanDelete") = 1)  and Session("dbKey") > 0 Then 
			If Session("dbConfirmDelete") > 0 Then %>
				<TD BGCOLOR=<%= bgcolor %> ALIGN="RIGHT"><A HREF="<%= Session("dbGenericPath")%>GenericDelete.asp?CMD='CON'&<%=strLink%>">Delete</A></TD>
<% 			Else %>
				<TD BGCOLOR=<%= bgcolor %> ALIGN="CENTER"><A HREF="<%= Session("dbGenericPath") %>GenericDelete.asp?<%= strLink %>">Delete</A></TD>
<%			End If %>
<%		End If %>
</TR>
<%	
	End If
	xrs.MoveNext
Loop
%>
</TABLE>
</TD></TR>
</TABLE>

<%
' Find out if there should be Backward or Forward Buttons on the table.
If 	intStartRec = 1 Then
	isPrev = False
Else
	isPrev = True
	PrevStart = intStartRec - intDisplayRecs
	If PrevStart < 1 Then PrevStart = 1
%>
<HR SIZE="1">
<FONT SIZE="+1"><STRONG><A HREF="GenericList.asp?START=<%=PrevStart%>">Previous</A></STRONG></FONT>
<%
End If
If NOT xrs.EOF Then
	NextStart = intStartRec + intDisplayRecs
	isMore = True
	If isPrev Then
		Response.Write "&nbsp;&nbsp;|&nbsp;&nbsp;"
	Else
		Response.Write "<HR SIZE=1>"
	End If
%>
<FONT SIZE="+1"><STRONG><A HREF="GenericList.asp?START=<%=NextStart%>">Next</A></STRONG></FONT>
<%
Else
	isMore = False
End If
%>
<HR SIZE="1">
<%
If intStopRec > intCount Then intStopRec = intCount
Response.Write "Records " & intStartRec & " to " & intStopRec
' Close recordset and connection
xrs.Close
Set xrs = Nothing
xConn.Close
Set xConn = Nothing
%>

<P>
<% If strSearchFields & "x" <> "x" Then %>
<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC>
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<FORM ACTION="GenericSearchResult.asp" METHOD="POST">
	<TR>
		<TD HEIGHT=0 BGCOLOR="WHITE" "ALIGN="LEFT">
			<FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
			Search For </TD>
		<TD HEIGHT=0 BGCOLOR="WHITE" ALIGN="LEFT">
			<INPUT TYPE="Text" NAME="strSearch" SIZE=40>
		</TD>
		<TD HEIGHT=0 BGCOLOR="WHITE" ALIGN="LEFT">
			<INPUT TYPE="Submit" NAME="Submit" VALUE="Search Now">
		</TD>
	</TR>
</FORM>
</TABLE>
</TD></TR>
</TABLE>
<% End If %>

<!-- Main Body End -->

<!-- Footer -->
<% If Session("dbFooter") = 1 Then %>
<!--#include file="GenericFooter.inc"-->
<% End If %>
</BODY>
</HTML>
