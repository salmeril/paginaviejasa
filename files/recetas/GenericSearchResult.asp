<%
' Generic Database - Search, Build WHERE clause
' Author: Eli Robillard
' E-Mail: erobillard@ofifc.org
' Revision History:
'  19 Feb  99 - Forgot to add brackets while building WHERE clause. Fixed now.
'  12 Jan  99 - Added [brackets] around field names for Access, it removes them again for SQL.
'  30 Nov  98 - File created

' Get parameters
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")
pSearch = Request.Form("strSearch")
SearchFields = Session("dbSearchFields")

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
strWhere = ""

For x = 1 to intFieldCount
	If Mid(SearchFields, x, 1) = "1" AND (xrs.Fields(x-1).Type = 200) Then 
		strWhere = strWhere & " ([" & xrs.Fields(x-1).Name & "] LIKE '%" & pSearch & "%') OR" 
	End If
Next 
If Right(strWhere,2) = "OR" Then
	strWhere = Left(strWhere, Len(strWhere)-2)
End If
If strType = "SQL" Then
	' SQL databases do not allow spaces or brackets in table or field names
	strWhere = Replace(strsql,"[","")
	strWhere = Replace(strsql,"]","")
End If
Session("dbWhere") = strWhere

xrs.Close
Set xrs = Nothing

Session("dbStartRec") = 1
Response.Redirect Session("dbGenericPath") & "GenericList.asp"
%>
