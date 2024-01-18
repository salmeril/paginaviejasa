<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Product Maintenance - File Browser
' Product  : CandyPress eCommerce Administration
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
const adminLevel = 1
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<%
'Work Fields
dim I, errMsg, upd
dim path, pathRel, delFile
dim viewLink, selectLink, deleteLink
dim fs, folder, item

'Database
dim mySQL, cn, rs, rs2

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Close Database Connection
call closeDB()

%>
<html>
<head>
	<title>File Browser</title>
	<style type="text/css">
	<!--
	BODY, B, TD, P {COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt;}
	-->
	</style>
</head>
<body>

<script language="Javascript">
<!--
	function confirmDelete()
	{
		var agree=confirm("This action can not be undone. Are you sure you want to continue?");
		if (agree)
			return true ;
		else
			return false ;
	}
-->
</script>

<P align=center>
	<b><font size=3>File Browser</font></b>
	<br><br>
</P>

<%
'Get path to requested folder
upd = Request.QueryString("upd")
if upd = "SI" or upd = "LI" then	'Images
	pathRel = pImagesDir
	path    = Server.MapPath(pImagesDir)
else								'Downloads
	pathRel = pDownloadDir
	path    = Server.MapPath(pDownloadDir)
end if

'Get delete file (if any)
delFile = trim(Request.QueryString("del"))
if demoMode = "Y" and len(delFile) > 0 then
	errMsg = "File Deletion is not allowed in Demo Mode."
	call endOfPage()
end if

on error resume next

'Get FileSystem Object
set fs = CreateObject("Scripting.FileSystemObject")
if err.number <> 0 then
	errMsg = err.Description
	call endOfPage()
end if

'If delete was requested, delete the file
if len(delFile) > 0 then
	fs.DeleteFile(path & "\" & delFile)
	if err.Number <> 0 then
		errMsg = err.Description & " : " & path & "\" & delFile
		call endOfPage()
	end if
end if
	
'Get Folder Object
set folder = fs.GetFolder(path)
if err.number <> 0 then
	errMsg = err.Description
	call endOfPage()
end if
		
'Get Recordset Object and Disconnect
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = adUseClient 
rs.CursorType = adOpenDynamic 
rs.Fields.Append "filename",200,255
rs.Fields.Append "filedate",200,255
rs.Fields.Append "filesize",200,255
rs.open
if err.number <> 0 then
	errMsg = err.Description
	call endOfPage()
end if

'Populate recordset and sort
for each item in folder.Files 
	rs.AddNew 
	rs.Fields("filename") = item.Name 
	rs.Fields("filedate") = item.dateLastModified
	rs.Fields("filesize") = item.Size
	rs.update 
next 
rs.Sort = "filename ASC" 
rs.MoveFirst

'Display files
Response.Write "<table border=0 cellspacing=0 cellpadding=1 width=""100%"" align=center>"
Response.Write "<tr><td nowrap colspan=4><span style=""FONT-FAMILY: courier; FONT-SIZE: 12pt;"">" & pathRel & "</span></center></td></tr>"
Response.Write "<tr><td bgcolor=#DDDDDD><b>Action</b></td><td bgcolor=#DDDDDD><b>Filename</b></td><td bgcolor=#DDDDDD><b>Modified</b></td><td bgcolor=#DDDDDD><b>Bytes</b></td></tr>"
do while not rs.EOF
				
	'Increment counter
	I = I + 1
					
	'Determine what "select"" link to show
	if upd = "SI" then		'Small Image
		selectLink = "<a href=""#"" onClick=""opener.document.prodForm.smallImageURL.value='" & rs("filename") & "'; window.close(); return false;""><img src=""x_edit.gif"" border=0 alt=Select></a>"
	elseif upd = "LI" then	'Large Image
		selectLink = "<a href=""#"" onClick=""opener.document.prodForm.imageURL.value='"      & rs("filename") & "'; window.close(); return false;""><img src=""x_edit.gif"" border=0 alt=Select></a>"
	else					'Downloads
		selectLink = "<a href=""#"" onClick=""opener.document.prodForm.fileName.value='"      & rs("filename") & "'; window.close(); return false;""><img src=""x_edit.gif"" border=0 alt=Select></a>"
	end if
					
	'Determine what "view" link to show
	if inStrRev(lCase(rs("filename")),".gif")  _
	or inStrRev(lCase(rs("filename")),".jpg")  _
	or inStrRev(lCase(rs("filename")),".jpeg") _
	or inStrRev(lCase(rs("filename")),".bmp") then
		viewLink = "<a href=""#"" onClick=""img" & I & ".src='" & pathRel & rs("filename") & "'; return false;""><img src=""x_view.gif"" border=0 alt=View></a>"
	else
		viewLink = "<img src=""x_cleardot.gif"" border=0 width=16 height=16>"
	end if
				
	'Determine what "delete" link to show
	deleteLink = "<a href=""SA_prodImg.asp?upd=" & upd & "&del=" & server.URLEncode(rs("filename")) & """ onClick=""return confirmDelete()"" ><img src=""x_delete.gif"" border=0 alt=Delete></a>"
					
	'Write image row
	Response.Write "" _
		& "<tr>" _
		& "  <td nowrap valign=middle> " _
		&      selectLink & " " _
		&      viewLink   & " " _
		&      deleteLink & " " _
		& "  </td>" _
		& "  <td nowrap>" & rs("filename") & "</td>" _
		& "  <td nowrap>" & rs("filedate") & "</td>" _
		& "  <td nowrap>" & rs("filesize") & "</td>" _
		& "</tr>" & vbCrlf _
		& "<tr>" _
		& "  <td colspan=4><img src=""x_cleardot.gif"" name=img" & I & "></td>" _
		& "</tr>" & vbCrlf
						
	'Next record
	rs.MoveNext
loop
Response.Write "</table>"
	
on error goto 0

call endOfPage()

'*********************************************************************
'End of page processing
'*********************************************************************
sub endOfPage()

	'Clean up
	set folder = nothing
	set fs     = nothing
	set rs	   = nothing

	'Check for error message
	if len(errMsg) <> 0 then
%>
		<p align=center>
			<span style="COLOR: red; FONT-SIZE: 10pt;">
				Error : <%=errMsg%>
			</span>
		</p>
		<p align=center>
			<a href="javascript:history.go(-1)">BACK</a>
		</p>
<%
	end if
%>
	</body></html>
<%
	Response.End
end sub
%>
