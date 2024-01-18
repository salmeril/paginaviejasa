<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Database functions.
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
	if UCase(dbLocked) = "Y" then
		call errorDB("<b>" & langErrStoreClosed & "</b>","")
	end if
	on error resume next
	set connTemp = server.createobject("adodb.connection")
	connTemp.Open connString
	if err.number <> 0 then
		dim errMsg
		errMsg = "" _
			& "<b>Number :</b> " & err.number & "<br><br>" _
			& "<b>Page :</b> "   & Request.ServerVariables("PATH_INFO") & "<br><br>" _
			& "<b>Desc :</b> "   & err.Description
		call errorDB("",errMsg)
	end if
	on error goto 0
end function

'*************************************************************************
'Close Database Connection
'*************************************************************************
function closeDB()
	on error resume next
	connTemp.close
	set connTemp = nothing
	on error goto 0
end function

'*************************************************************************
'Open RecordSet using "Execute" method
'*************************************************************************
function openRSexecute(mySQL)
	on error resume next
	set openRSexecute = conntemp.execute(mySQL)
	if err.number <> 0 then
		dim errMsg
		errMsg = "" _
			& "<b>Number :</b> " & err.number & "<br><br>" _
			& "<b>Page :</b> "   & Request.ServerVariables("PATH_INFO") & "<br><br>" _
			& "<b>Desc :</b> "   & err.Description & "<br><br>" _
			& "<b>SQL :</b> "    & mySQL
		call errorDB("",errMsg)
	end if
	on error goto 0
end function

'*************************************************************************
'Open RecordSet using "Open" method
'*************************************************************************
function openRSopen(dbSource,dbCursorLoc,dbCursorType,dbLockType,dbOptions,dbCache)
	on error resume next
	set openRSopen = Server.CreateObject("ADODB.Recordset")
	if dbCache > 0 then
		openRSopen.CacheSize = dbCache
	end if
	if dbCursorLoc > 0 then
		openRSopen.CursorLocation = dbCursorLoc
	end if
	openRSopen.Open dbSource,connTemp,dbCursorType,dbLockType,dbOptions
	if err.number <> 0 then
		dim errMsg
		errMsg = "" _
			& "<b>Number :</b> " & err.number & "<br><br>" _
			& "<b>Page :</b> "   & Request.ServerVariables("PATH_INFO") & "<br><br>" _
			& "<b>Desc :</b> "   & err.Description & "<br><br>" _
			& "<b>SQL :</b> "    & dbSource
		call errorDB("",errMsg)
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

'*************************************************************************
'Handle database errors
'*************************************************************************
sub errorDB(errMsgShow,errMsgHide)

	'Clear output buffer and declare work variables
	Response.Clear
	dim errMsg
	dim hideError
	
	'Decide which error to display, and if we must hide the error
	if len(trim(errMsgShow)) > 0 then
		errMsg		= trim(errMsgShow)
		hideError	= false
	else
		errMsg		= trim(errMsgHide)
		hideError	= true
	end if
	
	'Force detailed error to be displayed if debug mode is on
	on error resume next
	if UCase(debugMode) = "Y" then
		if err.number = 0 then
			hideError = false
		end if
	end if
	on error goto 0
%>
	<HTML>
	<HEAD></HEAD>
	<BODY>
		<P align=center>
			<br><br><br>
			<font face="verdana,arial" size="2" color=red>
				<b>System Error</b>
			</font><br><br>
			<table border="1" bgcolor="#EEEEEE" cellpadding="15" width="50%"><tr><td align=left>
				<font face="verdana,arial" size="2">
<%
					if hideError then
%>
						Note : The detail of this error can be 
						viewed by activating debug mode for 
						this store.
<%
					else
						Response.Write errMsg
					end if
%>
				</font>
			</td></tr></table>
		</P>
	</BODY>
	</HTML>
<%
	'Close open database connections and end
	call closeDB()
	Response.End
end sub
%>