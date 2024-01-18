<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Execute SQL statement
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
Option explicit
Response.Buffer = true
const adminLevel = 0
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<%
'Declare variables
dim mySQL, cn, rs
dim field, count

'*************************************************************************

'Are we in demo mode?
if demoMode = "Y" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("DEMO MODE. Sorry, this featured is NOT available in Demo Mode.")
end if

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Execute SQL Command</font></b>
	<br><br>
</P>

<%
mySql = trim(request.form("mySQL"))
if len(mySQL) = 0 then
	mySql = request.querystring("mySQL")
end if
%>

<b>Query: --> </b><font color=red><%=mySQL%></font>
<br><br>

<%

'Open DB
call openDb()

'Open RecordSet
set rs = openRSexecute(mySQL)

'Reset Counter
count = 1

if rs.State = adStateClosed then
	Response.Write "<br>Command Executed."
	
else

	if rs.eof then
		Response.Write "<br>Command Executed."
		
	else
		'Iterate through RecordSet
		Response.Write "<table border=1>"
		do while not rs.eof
			Response.Write "<tr>"
			for each field in rs.fields
				Response.Write "<td>"
				if count = 1 then
					response.write "<b>" & field.name & "</b>"
				else
					if isNull(rs(field.name)) then
						response.write "&lt;null&gt;&nbsp;"
					else
						response.write rs(field.name) & "&nbsp;"
					end if
				end if	 
				Response.Write "</td>"
			next
			Response.Write "</tr>"
			if count > 1 then
				rs.movenext
			end if
			count = count + 1
		loop
		Response.Write "</table>"
	end if

	'Close RecordSet
	call closeRS(rs)

end if

call closedb()

%>
<!--#include file="_INCfooter_.asp"-->
