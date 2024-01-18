<% 
' Generic Database - Exit 
' Author: Eli Robillard
' E-Mail: erobillard@ofifc.org
' URL: http://www.ofifc.org/Eli/ASP/GenericArticle.asp
' Revision History:
'  19 Feb  99 - Support for dbFooter setting
'  30 Nov  98 - File Created
'			  - Support for sub-tables

' Quick security check, make sure we have an active session
If Session("dbDispList") = "" or Session("dbConn") = "" Then 
	Response.Redirect Session("dbExitPage")
End If

' Get the key value of the record to display	
If Request.QueryString("KEY").Count > 0 Then
	dbGoSub = True
	strKey = Request.QueryString("KEY")
	Session("dbcurKey") = strKey
End If

' Get the parameters set in the Config File
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")

' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn

' Open Recordset and get the field count
strsql = "SELECT * FROM [" & strTable & "]"
If strType = "SQL" Then
	' SQL databases do not allow spaces or brackets in table or field names
	strsql = Replace(strsql,"[","")
	strsql = Replace(strsql,"]","")
End If
set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn
intFieldCount = xrs.Fields.Count
xrs.Close
Set xrs = Nothing
xConn.Close
Set xConn = Nothing

' Reset the Parameters
For x = 1 to intFieldCount
	Session("dbCombo" & x) = Null
	Session("dbDefault" & x) = Null
	Session("dbURLfor" & x) = Null
Next

Session("dbStartRec") = 1
Session("dbRecsPerPage") = 0
Session("dbFooter") = 0
Session("dbDispList") = ""
Session("dbDispView") = ""
Session("dbDispEdit") = ""
Session("dbSearchFields") = ""

If dbGoSub AND (Session("dbSubTable") & "x" <> "x") Then
	' If going to a sub-table
	arrSubTable = Split(Session("dbSubTable"),",")
	' Copy the subtable vals into another var and clear it so it doesn't think there's more below this one.
	Session("dbSubTableCopy") = Session("dbSubTable")
	Session("dbSubTable") = ""
	Session("dbIsSubTable") = True
	' Set the relation to the subtable
	Session("dbWhere") = QUOTE & arrSubTable(2) & " = " & strKey & QUOTE 
	strURL = arrSubTable(1)
	Response.Redirect strURL
Else
	' If exiting GenericDB
	Session("dbIsSubTable") = False
	Session("dbSubTableCopy") = ""
	Response.Redirect Session("dbExitPage")
End If
%>
