<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Functions used to Open and Close a database connection.
'          : Functions used to Open and Close a RecordSet.
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************

'*************************************************************************
'Declare some standard ADO variables
'*************************************************************************
Const adOpenKeyset		= 1
Const adOpenDynamic		= 2
Const adOpenStatic		= 3
Const adLockReadOnly	= 1
Const adLockOptimistic	= 3
Const adStateClosed		= &H00000000
Const adUseServer		= 2
Const adUseClient		= 3
Const adCmdText			= &H0001
Const adCmdTable		= &H0002

'*************************************************************************
'Open Database Connection
'*************************************************************************
function openDB()
	set cn = server.createobject("adodb.connection")
	on error resume next
	cn.Open ConnString
	if err.number <> 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(err.Description)
	end if
	on error goto 0
end function

'*************************************************************************
'Close Database Connection
'*************************************************************************
function closeDB()
	on error resume next
	cn.close
	set cn = nothing
	on error goto 0
end function

'*************************************************************************
'Open RecordSet using "Execute" method
'*************************************************************************
function openRSexecute(mySQL)
	on error resume next
	set openRSexecute = cn.execute(mySQL)
	if err.number <> 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(err.Description)
	end if
	on error goto 0
end function

'*************************************************************************
'Open RecordSet using "Open" method
'*************************************************************************
function openRSopen(dbSource,dbCursorLoc,dbCursorType,dbLockType,dbOptions,dbCache)
	set openRSopen = Server.CreateObject("ADODB.Recordset")
	if dbCache > 0 then
		openRSopen.CacheSize = dbCache
	end if
	if dbCursorLoc > 0 then
		openRSopen.CursorLocation = dbCursorLoc
	end if
	on error resume next
	openRSopen.Open dbSource,cn,dbCursorType,dbLockType,dbOptions
	if err.number <> 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(err.Description)
	end if
	on error goto 0
end function

'*************************************************************************
'Close Recordset
'*************************************************************************
function closeRS(rs)
	on error resume next
	rs.Close
	set rs = nothing
	on error goto 0
end function
%>