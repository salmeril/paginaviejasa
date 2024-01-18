<% 
' Generic Database - Exit 
' Notice: (c) 1998, 1999 Eli Robillard, All Rights Reserved. 
' E-Mail: erobillard@ofifc.org
' URL: http://www.ofifc.org/Eli/ASP/
' Revision History:
' 14 Jul 1999 - Added Response.Clear before Redirect for boneheaded MSIE browsers
'  6 Jul 1999 - Fix for subtables as suggested by Paul Reith
' 30 Jun 1999 - dbRequiredFields, dbUpdateFieldX, dbHaving
'  5 May 1999 - Fixed redirect problem for time-outs. Now redirect to GenericError.
'  1 Mar 1999 - Support for dbFont and dbFontSize settings
' 23 Feb 1999 - Added an option to reset rather than exit (redirect back to Lister).
' 19 Feb 1999 - Support for dbFooter 
' 30 Nov 1998 - First created or released

Response.Buffer = True
' Quick security check, make sure we have an active session
If (Session("dbDispList") & "x" = "x") or (Session("dbConn") & "x"  = "x") Then 
	Response.Clear
	Response.Redirect "GenericError.asp"
End If

doGoSub = False	' True if exiting to a subtable
doReset = False ' True if resetting to values from the config file

' Get the key value of the record to display
If Request.QueryString("KEY").Count > 0 Then
	dbGoSub = True
	strKey = Request.QueryString("KEY")
	' Suggested by Paul Reith:
	subkey=Request.querystring("KEY")
	Session("dbsubkey") = subkey
	Session("dbcurKey") = strKey
End If

' See if this is a reset or an exit
If Request.QueryString("CMD").Count > 0 Then
	' Reset is the only parameter right now
	strCmd = Request.QueryString("CMD")
	doReset = True
End If

' Get the parameters set in the Config File
strType = UCase(Session("dbType"))
strConn = Session("dbConn")
strTable = Session("dbRs")

' Open Recordset and get the field count
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn
strsql = "SELECT * FROM [" & strTable & "]"
Select Case strType
	Case "UDF" 
		strsql = "SELECT * FROM " & strTable
	Case "SQL" 
		strsql = Replace(strsql,"[","")
		strsql = Replace(strsql,"]","")
End Select
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
	Session("dbEMailfor" & x) = Null
	Session("dbUpdateField" & x) = Null
Next

' Zero all settings
Session("dbTitle") = ""
Session("dbCanAdd") = 0
Session("dbCanEdit") = 0
Session("dbCanDelete") = 0
Session("dbConfirmDelete") = 0
Session("dbState") = 1
Session("dbStartRec") = 1
Session("dbRecsPerPage") = 0
Session("dbFontSize") = 2
Session("dbFooter") = 0
Session("dbOrder") = 0
Session("dbDispList") = ""
Session("dbDispView") = ""
Session("dbDispEdit") = ""
Session("dbSearchFields") = ""
Session("dbRequiredFields") = ""
Session("dbTotalFields") = ""
Session("dbFields") = ""
Session("dbGroupBy") = ""
Session("dbHaving") = ""
Session("dbOrderBy") = ""
Session("dbWhereOld") = ""
Session("dbFieldNames") = ""
Session("dbFont") = ""
Session("dbBorderColor") = ""
Session("dbMenuColor") = ""
Session("dbEditTemplate") = ""
Session("dbViewTemplate") = ""
Session("dbBackText") = ""
Session("dbAddExtra") = ""

If dbGoSub AND (Session("dbSubTable") & "x" <> "x") Then
	' If going to a sub-table
	arrSubTable = Split(Session("dbSubTable"),",")
	' Copy the dbSubTable parm and clear it so it doesn't think there are more below this one.
	Session("dbSubTableCopy") = Session("dbSubTable")
	Session("dbSubTable") = ""
	Session("dbIsSubTable") = True
	' Set the relation to the subtable
	Session("dbWhere") = QUOTE & arrSubTable(2) & " = " & strKey & QUOTE 
	strURL = arrSubTable(1)
	Response.Clear
	Response.Redirect strURL
Else
	If doReset Then ' reread the config file
		Response.Clear
		Response.Redirect Session("dbViewPage")
	Else ' exit GenericDB
		Session("dbIsSubTable") = False
		Session("dbSubTableCopy") = ""
		Session("dbLastRs") = ""
		Response.Clear
		Response.Redirect Session("dbExitPage")
	End If
End If
%>
