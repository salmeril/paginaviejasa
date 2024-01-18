<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Upload Files
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
const adminLevel = 1
server.ScriptTimeout = 600 '10 Minutes
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<%
'Work Fields
dim strFile
Dim Uploader
dim File
dim fso
dim uplFolder
dim logonWinUser
dim maxFileSize

'Database
dim mySQL, cn, rs

'*************************************************************************

'Are we in test mode?
if demoMode = "Y" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("DEMO MODE. Sorry, this featured is NOT available in Demo Mode.")
end if

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Close Database Connection
call closedb()

'Set Maximum File Size - Will be user modifiable at a later date
maxFileSize = 0  '0 = Any file size

'Create a FileUploader object
Set Uploader = New FileUploader

'Initialize and start the upload of the files and form variables
on error resume next
Uploader.Upload()
if err.number = 0 then

	'If the user requested to logon to his windows account before 
	'attemtping the upload, run this first. Note that we have to do 
	'this AFTER calling Upload() above because the form variables are 
	'not available via 'request.form' due to the encoding being used.
	if UCase(Uploader.Form("logonWin")) = "Y" then
		logonWinUser = Request.ServerVariables("LOGON_USER")
		If IsEmpty(logonWinUser) Or IsNull(logonWinUser) Or logonWinUser="" Then
			Response.Status = "401 Access Denied"
			Response.Addheader "WWW-Authenticate", "BASIC"
			Response.Write "Invalid Logon. Please try again."
			Response.End
		end if
	end if

	'Check that at least one file was uploaded
	if not (Uploader.Files.Exists("file01") or Uploader.Files.Exists("file02") or Uploader.Files.Exists("file03") or Uploader.Files.Exists("file04") or Uploader.Files.Exists("file05")) then
		Response.Redirect "upload.asp?msg=" _
			& server.URLEncode("ERROR : At least one file must be selected.")
	end if
	
	'Check that a destination folder was selected
	uplFolder = trim(Uploader.Form("uplFolder"))
	if len(uplFolder) = 0 then
		Response.Redirect "upload.asp?msg=" _
			& server.URLEncode("ERROR : Invalid destination folder selected.")
	end if

	'File01
	If Uploader.Files.Exists("file01") Then
		call checkFileSize("file01")
		Uploader.Files("file01").SaveToDisk Server.MapPath(uplFolder)
		if err.number <> 0 then
			response.redirect "upload.asp?msg=" _
				& server.URLEncode("ERROR : File [" _
				& Uploader.Files("file01").FileName _
				& "] could not be saved to disk (" _
				& err.Description & ").")
		end if
	end If
	
	'File02
	If Uploader.Files.Exists("file02") Then
		call checkFileSize("file02")
		Uploader.Files("file02").SaveToDisk Server.MapPath(uplFolder)
		if err.number <> 0 then
			response.redirect "upload.asp?msg=" _
				& server.URLEncode("ERROR : File [" _
				& Uploader.Files("file02").FileName _
				& "] could not be saved to disk (" _
				& err.Description & ").")
		end if
	end If

	'File03
	If Uploader.Files.Exists("file03") Then
		call checkFileSize("file03")
		Uploader.Files("file03").SaveToDisk Server.MapPath(uplFolder)
		if err.number <> 0 then
			response.redirect "upload.asp?msg=" _
				& server.URLEncode("ERROR : File [" _
				& Uploader.Files("file03").FileName _
				& "] could not be saved to disk (" _
				& err.Description & ").")
		end if
	end If

	'File04
	If Uploader.Files.Exists("file04") Then
		call checkFileSize("file04")
		Uploader.Files("file04").SaveToDisk Server.MapPath(uplFolder)
		if err.number <> 0 then
			response.redirect "upload.asp?msg=" _
				& server.URLEncode("ERROR : File [" _
				& Uploader.Files("file04").FileName _
				& "] could not be saved to disk (" _
				& err.Description & ").")
		end if
	end If

	'File05
	If Uploader.Files.Exists("file05") Then
		call checkFileSize("file05")
		Uploader.Files("file05").SaveToDisk Server.MapPath(uplFolder)
		if err.number <> 0 then
			response.redirect "upload.asp?msg=" _
				& server.URLEncode("ERROR : File [" _
				& Uploader.Files("file05").FileName _
				& "] could not be saved to disk (" _
				& err.Description & ").")
		end if
	end If

'There was a problem	
else
	response.redirect "upload.asp?msg=" _
		& server.URLEncode("ERROR : Upload Object could not process request (" _
		& err.Description & ").")
end if

'If we get this far, everything was OK
Response.Redirect "upload.asp?msg=" & server.URLEncode("<font color=green>SUCCESS : Files were uploaded.</font>")

'**********************************************************************
'Check File Size
'**********************************************************************
sub checkFileSize(strFile)
	'Has a maximum size been specified?
	if maxFileSize > 0 then
		if Uploader.Files(strFile).FileSize > maxFileSize then
			response.redirect "upload.asp?msg=" & server.URLEncode("ERROR : Maximum File Size is " & maxFileSize & " bytes.")
		end if
	end if
end sub

'**********************************************************************
'File Upload Classes
'**********************************************************************
Class FileUploader
	Public  Files
	Private mcolFormElem

	Private Sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")
		Set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(Files) Then
			Files.RemoveAll()
			Set Files = Nothing
		End If
		If IsObject(mcolFormElem) Then
			mcolFormElem.RemoveAll()
			Set mcolFormElem = Nothing
		End If
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If mcolFormElem.Exists(LCase(sIndex)) Then Form = mcolFormElem.Item(LCase(sIndex))
	End Property

	Public Default Sub Upload()
		Dim biData, sInputName
		Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos, nTempPos
		Dim nPosFile, nPosBound

		biData = Request.BinaryRead(Request.TotalBytes)
		nPosBegin = 1
		nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
		If (nPosEnd-nPosBegin) <= 0 Then Exit Sub

		vDataBounds = MidB(biData, nPosBegin, nPosEnd - nPosBegin)
		nDataBoundPos = InstrB(1, biData, vDataBounds)

		Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))

			nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InstrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InstrB(nPosEnd, biData, vDataBounds)

			If nPosFile <> 0 And nPosFile < nPosBound Then
				Dim oUploadFile, sFileName
				Set oUploadFile = New UploadedFile

				nPosBegin = nPosFile + 10
				nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
				sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))

				nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
				nPosBegin = nPos + 14
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
				nTempPos = InStrB(nPosBegin, biData, CByteString(";"))
				If nTempPos < nPosEnd And nTempPos > 0 Then nPosEnd = nTempPos

				oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				nPosBegin = InStrB(nPosBegin, biData, CByteString(Chr(13))) + 4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
				
				If oUploadFile.FileSize > 0 And Not oUploadFile.FileName = "" Then Files.Add sInputName, oUploadFile
			Else
				nPos = InstrB(nPos, biData, CByteString(Chr(13)))
				nPosBegin = nPos + 4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				If Not mcolFormElem.Exists(LCase(sInputName)) Then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin)) Else mcolFormElem(LCase(sInputName)) = mcolFormElem(LCase(sInputName)) & ", " & CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			End If

			nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
		Loop
	End Sub

	'String to byte string conversion
	Private Function CByteString(sString)
		Dim nIndex
		For nIndex = 1 to Len(sString)
		   CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
		Next
	End Function

	'Byte string to string conversion
	Private Function CWideString(bsString)
		Dim nIndex
		CWideString =""
		For nIndex = 1 to LenB(bsString)
		   CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
		Next
	End Function
End Class

Class UploadedFile
	Public ContentType
	Public FileName
	Public FileData
	
	Public Property Get FileSize()
		FileSize = LenB(FileData)
	End Property

	Public Sub SaveToDisk(sPath)
		Dim oFS, oFile
		Dim nIndex
	
		If sPath = "" Or FileName = "" Then Exit Sub
		If Mid(sPath, Len(sPath)) <> "\" Then sPath = sPath & "\"
	
		Set oFS = Server.CreateObject("Scripting.FileSystemObject")
		If Not oFS.FolderExists(sPath) Then Exit Sub
		
		Set oFile = oFS.CreateTextFile(sPath & FileName, True)
		
		For nIndex = 1 to LenB(FileData)
		    oFile.Write Chr(AscB(MidB(FileData,nIndex,1)))
		Next

		oFile.Close
	End Sub
	
	Public Sub SaveToDatabase(ByRef oField)
		If LenB(FileData) = 0 Then Exit Sub
		
		If IsObject(oField) Then
			oField.AppendChunk FileData
		End If
	End Sub

End Class
%>