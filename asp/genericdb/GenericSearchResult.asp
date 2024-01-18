<%
' Generic Database - Search, Build WHERE clause
' Notice: (c) 1998, 1999 Eli Robillard, All Rights Reserved. 
' E-Mail: erobillard@ofifc.org
' URL: http://www.ofifc.org/Eli/ASP/
' Revision History:
' 15 Jun 1999 - Search Memo fields too
'				Fix for SQL server to remove the final AND in certain cases.
' 30 Nov 1998 - File created

' Prevent caching
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Cache-Control", "must-revalidate"
Response.AddHeader "Cache-Control", "no-cache"

' Get parameters
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")
strFields = Session("dbFields")
strGroupBy = Session("dbGroupBy")
pSearch = Request.Form("strSearch")
SearchFields = Session("dbSearchFields")

' Don't allow a one character search
If Len(Trim(pSearch)) < 2 Then Response.Redirect Session("dbGenericPath") & "GenericList.asp"

' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn

' Open Recordset and get the field info
strsql = "SELECT " & strFields & " FROM [" & strTable & "]"
Select Case strType
	Case "UDF" 
		strsql = "SELECT " & strFields & " FROM " & strTable
	Case "SQL" 
		strsql = Replace(strsql,"[","")
		strsql = Replace(strsql,"]","")
End Select
If Not Trim(strGroupBy) = "" Then
	strsql = strsql & " GROUP BY " & strGroupBy
	intAllowSort = 0
End If	
set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn
intFieldCount = xrs.Fields.Count
strWhere = ""

For x = 1 to intFieldCount
	If Mid(SearchFields, x, 1) = "1" AND (xrs.Fields(x-1).Type >= 200) AND (xrs.Fields(x-1).Type <= 203) Then strWhere = strWhere & " ([" & xrs.Fields(x-1).Name & "] LIKE '%" & pSearch & "%') OR" 
Next 
If Right(strWhere,2) = "OR" Then strWhere = Left(strWhere, Len(strWhere)-2)
If strType = "SQL" Then
' Strip brackets for SQL
	strWhere = Replace(strWhere,"[","")
	strWhere = Replace(strWhere,"]","")
End If

' If the where clause was defined in the config file, set the state, otherwise proceed normally.
If (Trim(Session("dbWhere")) & "x" = "x") OR Session("dbState") = 2 Then
	' If a where clause was not set in the config file, then ad hoc searches are allowed
	' dbState= 1: Normal Operation; 2: Search, no previous Where, 3: Search, Where set in config
	Session("dbState") = 2
	Session("dbWhere") = strWhere
Else
	Session("dbState") = 3
	if strWhere & "x" = "x" Then
		Session("dbWhere") = "(" & Session("dbWhere") & ")"
	else
		Session("dbWhere") = "(" & Session("dbWhere") & ") AND (" & strWhere & ")"
	end if
End If
xrs.Close
Set xrs = Nothing

Session("dbStartRec") = 1
Response.Redirect Session("dbGenericPath") & "GenericList.asp"
%>
