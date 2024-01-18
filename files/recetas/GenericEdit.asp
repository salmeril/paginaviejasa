<% 
' Generic Database - Edit Record
' Author: Eli Robillard
' E-Mail: erobillard@ofifc.org
' URL: http://www.ofifc.org/Eli/ASP/GenericArticle.asp
' Revision History:
'  18 Feb  99 - Addred redirect back to Lister after Update
'  15 Feb  99 - Changed format of title / button bar
'			  - Testing workaround FrontPage problems, suggested by Leif from lcom.dk
'			  - On adding a new record, using Dynamic Recordset (2) and Optimistic locking (3)
'				instead of (1,2)
'  12 Jan  99 - Fixed support for reals and currency (didn't set defaults properly before)
'  30 Nov  98 - Added 2 and 4-byte reals, currency to types supported
'			  - Combo box support, ie: Session("dbCombo2") = "LIST, Value 1, Desc 1, Value 2, Desc 2"
'			  - Default values support, ie: Session("dbDefault3") = Date()
'			  - Support for sub-tables
'  24 Nov  98 - Session("dbDispEdit") var. Like the dbDispList var, it allows you to set which
'				fields are displayed on the Edit screen.
'			  - Added BGCOLOR=#FFFFFF (white) to the title display, Explorer was overlapping
'				the title with the blue from the rightmost cell
'  27 Sept 98 - Added Session("dbType") var. If set to "SQL", the brackets are stripped out of
'				the SQL queries. 
'  18 Sept 98 - After an add, client is redirected to GenericList.asp. This prevents the problem
'				with redisplaying the record without being able to retrieve the new key value.
'			  - Removed rs.Find references (ADO 2.0) and recreated the functionality with SQL.
'  11 Sept 98 - Move the quick security checks into the Request.Querystring checks.
'				Before if you had dbCanAdd but not dbCanEdit rights an Add would fail.
' 	9 Sept 98 - Last modified

' Check for an active session
If Session("dbConn") = "" Then 
	Response.Redirect "GenericError.asp"
End If

' Get info from Session vars (kinda like parameters)
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")
strDisplay = Session("dbDispEdit")
strKeyField = Session("dbKey")
IsSubTable = Session("dbIsSubTable")

Session("dbCanEdit") = 1
SubmitValue = "Update"
Action = "GET"

If Request.QueryString("KEY").Count > 0 Then
	' Quick security check for Edit rights
	If Session("dbCanEdit") <> 1 Then
		Response.Redirect Session("dbViewPage")
	End If
	strKey = Request.QueryString("KEY")
	Session("dbcurKey") = strKey
	Action = "GET"
ElseIf Request.QueryString("CMD").Count > 0 Then
	' Quick security check for Add rights
	If Session("dbCanAdd") <> 1 Then
		Response.Redirect Session("dbViewPage")
	End If

	strCMD = Request.QueryString("CMD")
	If strCMD = "NEW" Then
		Action = "NEW"
	End If
Else
	Action = Left(UCase(Request.Form("Action")),3)
End If

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
ReDim aFields(intFieldCount,4)
For x = 1 to intFieldCount
	aFields(x, 1) = xrs.Fields(x-1).Name 
	aFields(x, 2) = xrs.Fields(x-1).Type 
	aFields(x, 3) = xrs.Fields(x-1).DefinedSize 
Next 
xrs.Close
Set xrs = Nothing

' Load the results of the last form view (GET or UPDATE)
For x = 1 to intFieldCount
	aFields(x,4) = Request.Form(aFields(x,1))
Next 

Select Case Action
	Case "ADD" ' Insert the new record into the database
		' Data validation 
		For x = 1 to intFieldCount
			Select Case aFields(x, 2) 
				Case 2, 3, 6	' 2 Byte Integer, 4 Byte Integer, Currency
					tFLD = aFields(x,4)
					xInt = 0
					If IsNumeric(tFLD) Then 
						aFields(x,4) = tFLD
					Else
						aFields(x,4) = xInt
					End If
				Case 11 	' Boolean True/False
					If aFields(x,4) = "Yes" Then
						aFields(x,4) = True
					Else
						aFields(x,4) = False
					End If
				Case 135	' Date / Time Stamp, usually created with the Now() function
					If (aFields(x,4) & "x" = "x") OR NOT IsDate(aFields(x,4)) Then
						aFields(x,4) = Null
					Else
						aFields(x,4) = CDate(aFields(x,4))
					End If
				Case 200, 201	' String or Memo
					tFLD = Trim(aFields(x,4))
					If tFLD & "x" = "x" Then
						tFLD = " "
					End If
					aFields(x,4) = tFLD
			End Select
		Next
		
		Set xrs =  Server.CreateObject("ADODB.Recordset")
		xrs.Open strTable, xConn, 2, 3
		xrs.AddNew
		
		For x = 1 to intFieldCount
			If strKeyField = x Then
				' Assume Key field is a counter and don't change it 
			Else
				xrs.Fields(x-1) = aFields(x,4)	
			End If
		Next 
		
		xrs.Update
		xrs.Close
		Set xrs = Nothing
		Response.Redirect Session("dbViewPage")
				
	Case "NEW": ' Load a blank form

		SUBMITVALUE = "Add"

		' Initialize Fields
		For x = 1 to intFieldCount
			If strKeyField = aFields(x,1) Then
				' Don't try to change the counter
			Else
				' Check if a default has been specified
				strDefault = "dbDefault" & x
				If Session(strDefault) & "x" = "x" Then
					Select Case aFields(x, 2) 
						Case 2, 3, 4, 5, 6
						' 2-Byte, 4-Byte Integer, Single, Double Precision Real, Currency
							aFields(x,4) = 0
						Case 11
						' Boolean True/False
							aFields(x,4) = "No"
						Case 135
						' Date / Time Stamp
							aFields(x,4) = ""
						Case 200
						' String
							aFields(x,4) = ""
						Case 201
						' Memo
							aFields(x,4) = ""
					End Select
				Else
					aFields(x,4) = Session(strDefault)
				End If
				If IsSubTable Then
					arrSubTable = Split(Session("dbSubTableCopy"),",")
					If Trim(aFields(x,1)) = Trim(arrSubTable(2)) Then
						aFields(x,4) = Session("dbcurKey")
					End If
				End If
			End If
		Next 

	Case "GET": ' Get a record to display
		strsql = "SELECT * FROM [" & strTable & "] WHERE [" & aFields(strKeyField,1) & "]=" & strKey
		If strType = "SQL" Then
			' SQL databases do not allow spaces or brackets in table or field names
			strsql = Replace(strsql,"[","")
			strsql = Replace(strsql,"]","")
		End If
		set xrs = Server.CreateObject("ADODB.Recordset")
		xrs.Open strsql, xConn
		xrs.MoveFirst
		If xrs.EOF Then
			Response.Redirect Session("dbViewPage")
		End If

		' Get the field contents
		For x = 1 to intFieldCount
			If aFields(x,2) = 11 Then
				If xrs(x-1) Then
					aFields(x,4) = "Yes"
				Else
					aFields(x,4) = "No"
				End If
			Else
				aFields(x,4) = xrs(x-1)
			End If
		Next 

		xrs.Close
		Set xrs = Nothing

	Case "UPD": ' Update
		' Open record
		strsql = "SELECT * FROM [" & strTable & "] WHERE [" & aFields(strKeyField,1) & "]=" & Session("dbcurKey")
		If strType = "SQL" Then
			' SQL databases do not allow spaces or brackets in table or field names
			strsql = Replace(strsql,"[","")
			strsql = Replace(strsql,"]","")
		End If
		set xrs = Server.CreateObject("ADODB.Recordset")
		xrs.Open strsql, xConn, 1, 2

		If xrs.EOF Then
			Response.Redirect Session("dbViewPage")
		End If

		For x = 1 to intFieldCount
			If strKeyField = x Then
				' Don't try to change the counter
			Else
				Select Case aFields(x,2) 
					Case 2
					' 2 Byte Integer
						tFLD = aFields(x,4)
						If IsNumeric(tFLD) Then 
							xrs(x-1) = CInt(tFLD)
						Else
							xrs(x-1) = 0
						End If
					Case 3
					' 4 Byte Integer
						tFLD = aFields(x,4)
						If IsNumeric(tFLD) Then 
							xrs(x-1) = CLng(tFLD)
						Else
							xrs(x-1) = 0
						End If
					Case 4
					' Single-Precision Floating Point
						tFLD = aFields(x,4)
						If IsNumeric(tFLD) Then 
							xrs(x-1) = CSng(tFLD)
						Else
							xrs(x-1) = 0
						End If
					Case 5, 6
					' Double-Precision Floating Point, Currency
						tFLD = aFields(x,4)
						If IsNumeric(tFLD) Then 
							xrs(x-1) = CDbl(tFLD)
						Else
							xrs(x-1) = 0
						End If
					Case 11
					' Boolean True/False
						tFLD = aFields(x,4)
						If tFLD = "Yes" Then
							xrs(x-1) = True
						Else
							xrs(x-1) = False
						End If
					Case 135
					' Date / Time Stamp, usually created with the Now() function
						If IsDate(aFields(x,4)) Then
							xrs(x-1) = CDate(aFields(x,4))
						Else
							' ERR: Bad date format
						End If
					Case 200, 201
					' String or Memo
						tFLD = Trim(aFields(x,4))
						If tFLD & "x" = "x" Then
							tFLD = " "
						End If
						xrs(x-1) = tFLD
				End Select
			End If
		Next 
		xrs.Update
		xrs.Close
		Set xrs = Nothing
		xConn.Close
		Set xConn = Nothing
		Response.Redirect Session("dbViewPage")
End Select
%>
<HTML>
<HEAD>
	<TITLE><%=Session("dbTitle")%> - Edit Mode</TITLE>
</HEAD>
<BODY>
<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC WIDTH="100%">
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<TR>
	<TD BGCOLOR=#99CCCC ALIGN="RIGHT" WIDTH="*"><FONT SIZE=3 FACE="Verdana, Arial, Helvetica">
		<A HREF="<%=Session("dbViewPage")%>">Back to List</A>
	</TD>
</TR>
<TR><TD ALIGN="RIGHT" BGCOLOR=#FFFFFF ROWSPAN=2 WIDTH="1"><FONT SIZE=5 FACE="Verdana, Arial, Helvetica"><STRONG><EM><%=Session("dbTitle")%> - Edit Mode</EM></STRONG></FONT> </TD></TR>
</TABLE></TD></TR></TABLE>

<FORM ACTION="GenericEdit.asp" METHOD=POST>
<INPUT TYPE="SUBMIT" NAME="Action" VALUE="<%=SUBMITVALUE%>">
<P>
<TABLE CELLPADDING=1 CELLSPACING=0 BORDER=0 BGCOLOR=#99CCCC>
<TR><TD>
<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 WIDTH=100% BGCOLOR=#99CCCC>
<% For x = 1 to intFieldCount 
	If (strKeyField = x) OR Mid(strDisplay, x, 1) = "0" Then
		' The typeval is included to fix a conflict with FrontPage, which doesn't like explicit declaration of "HIDDEN" here.
		typeVal = "Hidden" %>
		<INPUT TYPE="<%=typeVal%>" NAME="<%=aFields(x,1) %>" VALUE="<%=aFields(x,4) %>">
<% 	Else %>
		<TR BGCOLOR="#FFFFCC" ALIGN="LEFT">
		<TD >
			<% Response.Write aFields(x,1) %>
		</TD>

<% 		If aFields(x,1) = "Password" Then %>
			<TD BGCOLOR="White" ALIGN="LEFT">
			<INPUT TYPE="Password" NAME="<%=aFields(x,1)%>" VALUE="<%=aFields(x,4)%>" SIZE=40 MAXLENGTH="<%=aFields(x,3)%>"> </TD>

<%		Else 
			Select Case aFields(x,2) 
				Case 2 ' 2-Byte Integer %>
					<TD BGCOLOR="White" ALIGN="LEFT">
					<INPUT TYPE="Text" NAME="<%= aFields(x,1) %>" VALUE="<%= aFields(x,4) %>" SIZE=10 MAXLENGTH=4></TD>

				<% Case 3 ' 4-Byte Integer %>
					<TD BGCOLOR="White" ALIGN="LEFT">
					<INPUT TYPE="Text" NAME="<%= aFields(x,1) %>" VALUE="<%= aFields(x,4) %>" SIZE=10 MAXLENGTH=8></TD>
	
				<% Case 4, 5 ' Floating point %>
					<TD BGCOLOR="White" ALIGN="LEFT">
					<INPUT TYPE="Text" NAME="<%= aFields(x,1) %>" VALUE="<%= aFields(x,4) %>" SIZE=10 MAXLENGTH=8></TD>
	
				<% Case 6 ' Currency %>
					<TD BGCOLOR="White" ALIGN="LEFT">$
					<INPUT TYPE="Text" NAME="<%= aFields(x,1) %>" VALUE="<%= aFields(x,4) %>" SIZE=10 MAXLENGTH=8></TD>
	
				<% Case 11 	' Boolean True/False %>
					<TD BGCOLOR="White" ALIGN="LEFT">
					<INPUT TYPE="Radio" NAME="<%=aFields(x,1)%>" <% If aFields(x,4) = "Yes" Then %>CHECKED<% End If %> VALUE="Yes">Yes
						<INPUT TYPE="Radio" NAME="<%=aFields(x,1)%>" <% If aFields(x,4) = "No" Then %>CHECKED<% End If %> VALUE="No">No
					</TD>

				<% Case 135 ' Date / Time Stamp, usually created with the Now() function %>
					<TD BGCOLOR="White" ALIGN="LEFT">
					<INPUT TYPE="Text" NAME="<%= aFields(x,1) %>" VALUE="<%= aFields(x,4) %>" SIZE=40 MAXLENGTH=40></TD>

				<% Case 200 ' String %>
					<TD BGCOLOR="White" ALIGN="LEFT">
					<% 	strCombo = "dbCombo" & CStr(x)
						If Session(strCombo) & "x" = "x" Then %>
							<INPUT TYPE="Text" NAME="<%=aFields(x,1)%>" VALUE="<%=aFields(x,4)%>" SIZE=40 MAXLENGTH="<%=aFields(x,3)%>"> </TD>
<% 						Else %>
							<SELECT NAME="<%=aFields(x,1)%>" SIZE="1">
<%							arrCombo = Split(Session(strCombo),",")
							If UCase(arrCombo(0)) = "LIST" Then
								For y = 1 to UBound(arrCombo) Step 2
									arrCombo(y) = LTrim(arrCombo(y))
									arrCombo(y+1) = LTrim(arrCombo(y+1))
%>
									<OPTION VALUE="<%= arrCombo(y) %>" <% If arrCombo(y+1)=aFields(x,4) Then Response.Write" SELECTED" %>><%= arrCombo(y+1) %>
<%								Next
							Else
								' Read list from a table
							End If
%>
						</SELECT>
<% End If %>		
				<% Case 201 ' Memo %>
					<TD BGCOLOR="White" ALIGN="LEFT">
					<TEXTAREA NAME="<%=aFields(x,1)%>" COLS="50" ROWS="5" WRAP="VIRTUAL"><%=aFields(x,4)%></TEXTAREA></TD>
		
<% 			End Select 
 		End If %>
		</TR>
<% 	End If %>
<% Next %>
</TABLE>
</TD></TR>
</TABLE>
<P>
<INPUT TYPE="SUBMIT" NAME="Action" VALUE="<%=SUBMITVALUE%>">
</FORM>

<!-- Footer -->
<% If Session("dbFooter") = 1 Then %>
<!--#include file="GenericFooter.inc"-->
<% End If %>
</BODY>
</HTML>

