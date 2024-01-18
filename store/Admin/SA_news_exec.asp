<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Newsletters and Mailing List
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
server.ScriptTimeout = 18000 'Set to 5 hours due to newsletters
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="../Scripts/_INCappFunctions_.asp"-->
<!--#include file="../Scripts/_INCappEmail_.asp"-->
<%
'Declare variables
dim mySQL, cn, rs

'Newsletters
dim idNews
dim newsBookmark
dim newsSubj
dim newsBody
dim newsPreview
dim contType

'Work Fields
dim custType
dim custPaid
dim action
dim custList
dim I, I2

'Additional Newsletter Variables
dim custListEmail(20) 'Change to modify email batch size
dim strUA
dim nPctComplete

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Get general form values
custType	= trim(Request.Form("custType"))
custPaid	= trim(Request.Form("custPaid"))
action		= trim(Request.Form("action"))

'Get newsletter form values
idNews		= trim(Request.Form("idNews"))
newsBookmark= trim(Request.Form("newsBookmark"))
newsSubj	= trim(Request.Form("newsSubj"))
newsBody	= trim(Request.Form("newsBody"))
newsPreview = trim(Request.Form("newsPreview"))
contType    = trim(Request.Form("contType"))

'Check custType
if  custType <> "A" _
and custType <> "I" _
and custType <> "O" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Customer selection.")
end if

'Check custPaid
if custPaid <> "Y" then
	custPaid = "N"
end if

'Check Action
if  action <> "D" _
and action <> "F" _
and action <> "E" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Action selected.")
end if

'If newsletter, check that email is enabled
if action = "E" and mailComp = "0" then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Email must be enabled in order to send newsletters.")
end if

'Check newsletter info
if action = "E" then
	if len(idNews) > 0 then
		idNews = cLng(idNews)
	end if
	if len(newsSubj) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Subject for Newsletter.")
	end if
	if len(newsBody) = 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Invalid Message for Newsletter.")
	end if
	if newsPreview <> "Y" then
		newsPreview = "N"
	end if
	if isNumeric(contType) then
		contType = CLng(contType)
	else
		contType = 0
	end if
end if

'If the user requested a Preview newsletter, we can skip most of the 
'code and just execute the bit of code below.
if action = "E" and newsPreview = "Y" then
%>
	<!--#include file="_INCheader_.asp"-->
	<P align=left>
		<b><font size=3>Newsletters and Mailing Lists</font></b>
		<br><br>
	</P>
<%
	'Send emails
	on error resume next
	call sendmail (pCompany, pEmailSales, pEmailSales, newsSubj, newsBody, contType)
	if err.number <> 0 then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("An error occurred while sending email." & err.Description)
	end if
	on error goto 0
%>
	<span class="textBlockHead">Newsletter Preview</span><br>
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<tr><td>
		A preview of the newsletter was sent to :<br><br>
		
		<b><font color=red><%=pEmailSales%></font></b><br><br>
		
		Check your Email Inbox to see if this email was delivered, 
		and if the content and format of the email is correct. Press 
		the '<b>BACK</b>' button on your browser to return to the 
		newsletter to make changes and submit another preview, or to 
		submit the newsletter to all your customers (remember to 
		un-check the 'Preview' box first).
	</td></tr>
	</table>
	<!--#include file="_INCfooter_.asp"-->
<%
	'Close Database
	call closedb()

	'End script execution here
	Response.End
end if

'Construct SQL statement
'Notes : 
'1. Because the recordset is assigned to an array using getRows(), the 
'   position of the columns in the SQL statement is important as the 
'   rest of the code expects the columns to be located at pre-determined 
'   indexes. Email = index 0, lastName = index 1, name = index 2.
'2. The sort order is also important due to the fact that the sort 
'   order field (in this case Email) is used as a bookmark to restart 
'   from a particular position onwards.
'
'Specify columns and table
mySQL = "SELECT a.email, a.lastname, a.name " _
	  & "FROM   customer a " _
	  & "WHERE  1 = 1 "
'Opt-In customers
if custType = "I" then
	mySQL = mySQL & "AND a.futureMail = 'Y' "
end if
'Opt-Out customers
if custType = "O" then
	mySQL = mySQL & "AND (a.futureMail = 'N' OR a.futureMail = '' OR a.futureMail IS NULL) "
end if
'Customers who have PAID orders
if custPaid = "Y" then
	mySQL = mySQL & "AND EXISTS (SELECT b.idOrder FROM cartHead b WHERE b.idCust = a.idCust AND (b.orderStatus = '1' OR b.orderStatus = '2' OR b.orderStatus = '7')) "
end if
'Start from a certain email
if action = "E" and len(newsBookmark) > 0 then
	mySQL = mySQL & "AND email > '" & newsBookmark & "' "
end if
'Sort the recordset
mySQL = mySQL & "ORDER BY a.email "

'Execute SQL query and assign recordset to an array
set rs = openRSexecute(mySQL)
if rs.EOF then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("No customers were found matching your search criteria.")
else
	custList = rs.GetRows
end if
call closeRS(rs)

'If this is a newsletter, save the newsletter detail in the database
if action = "E" then

	'Check if we should UPDATE or INSERT newsletter info
	if len(idNews) > 0 and isNumeric(idNews) then
		mySQL = "UPDATE newsletters " _
			  & "SET    newsDate   = '" & now()						& "', " _
  			  & "       newsDateInt= '" & dateInt(now())			& "', " _
			  & "       newsSubj   = '" & replace(newsSubj,"'","''")& "', " _
			  & "       newsBody   = '" & replace(newsBody,"'","''")& "'  " _
			  & "WHERE  idNews = " & idNews
		set rs = openRSexecute(mySQL)
		call closeRS(rs)
	else
		set rs = openRSopen("newsletters",adUseServer,adOpenKeySet,adLockOptimistic,adCmdTable,0)
		rs.AddNew
		rs("newsDate")   = now()
		rs("newsDateInt")= dateInt(now())
		rs("newsSubj")   = newsSubj
		rs("newsBody")   = newsBody
		rs.Update
		idNews			 = rs("idNews") '@@identity
		call closeRS(rs)
	end if
	
end if

'Close Database
call closedb()

'DOWNLOAD customer list
if action = "F" then

	'Send headers for file name and content type changes
	Response.AddHeader "Content-Disposition", "attachment; filename=MailList.csv"
	Response.ContentType = "application/text"
	
	'Write results to output file
	for I = 0 to ubound(custList,2)
		Response.Write """" & custList(0,I) & """," & """" & custList(1,I) & "," & custList(2,I) & """" & vbCrLf
	next
	
	'End script execution
	Response.End
	
end if
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Newsletters and Mailing Lists</font></b>
	<br><br>
</P>

<%
'NEWSLETTERS via email
if action = "E" then

	'Determine user agent (browser)
	strUA = Request.ServerVariables("HTTP_USER_AGENT")
	If InStr(UCase(strUA), "MSIE") Then
		strUA = "IE"
	else
		strUA = "NS"
	end if
%>
	<!-- Create Progress Bar Pop-Up Window -->
	<SCRIPT LANGUAGE="JavaScript">
	<!-- 
		var progWin     = null;
		var progWinOpen = 0;
		progWin = window.open("","progWin","width=180,height=120,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,screenX=20,screenY=40,left=20,top=40");
	
		progWin.document.write('<html><head><title>Progress Bar</title></head>');	
		progWin.document.write('<body style="font-family: Verdana, Arial; font-size: 8pt;" onLoad="window.opener.progWinOpen=1" onUnload="window.opener.progWinOpen=0">');
		progWin.document.write('<center>');
		progWin.document.write('  <b style="color: #800000; background-color: #dddddd; padding: 2px;">Progress</b><br><br>');
<%
		if strUA = "IE" then
%>
		progWin.document.write('  <DIV STYLE="width:100px;height:16px;border-width:1px;border-style:solid;border-color:black" align="left">');
		progWin.document.write('    <DIV ID="progWinStatus" STYLE="width:0px;height:15px;background-color:red"></DIV>');
		progWin.document.write('  </DIV><br>');
<%
		else
%>
		progWin.document.write('  <form name=frmProgWin><b>Sent : </b><input type=text size=4 name=progWinStatus id=progWinStatus value=0></form>');
<%
		end if
%>
		progWin.document.write('  <b>Total To Send :</b> '+<%=ubound(custList,2)+1%>+'<br><br>');
		progWin.document.write('  <i>Please Wait ...</i>');
		progWin.document.write('</center>');
		progWin.document.write('</body>');
		progWin.document.write('</html>');
		
		progWin.document.close();
	//-->
	</SCRIPT>
<%
	'Send all buffered HTML to the browser. Without this, the user 
	'will wait until all processing has been completed before noticing 
	'a change in their browser, which may lead them to believe that 
	'nothing is happening. This will also pop-up the progress window.
	Response.Flush

	'Reset email array counter
	I2 = 0
%>
	<span class="textBlockHead">Newsletter Report</span><br>
	<table border=0 cellspacing=0 cellpadding=5 class="textBlock">
	<tr><td nowrap>
	
		<b>Total number of emails to send : <%=ubound(custList,2)+1%></b><br><br>
<%
		'Loop through recordset array. Emails are assigned to an email 
		'array (from the recordset array) so they can be sent in batches.
		for I = 0 to ubound(custList,2)
	
			'Move email address to email array
			custListEmail(I2) = custList(0,I)
			
			'Increment email array counter
			I2 = I2 + 1
			
			'Check if we need to send a batch of emails
			if I2 >= ubound(custListEmail) or I >= ubound(custList,2) then
			
				'Send a batch of emails
				on error resume next
				call sendmail (pCompany, pEmailSales, custListEmail, newsSubj, newsBody, contType)
				if err.number <> 0 then
					response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("An error occurred while sending email." & err.Description)
				end if
				on error goto 0
				
				'Update checkpoint. We open and close the database 
				'connection before and after each update because the 
				'script may run for a very long time, so we need to 
				'ensure that other scripts also get an opportunity.
				call openDb()
				if I < ubound(custList,2) then
					mySQL = "UPDATE newsletters " _
						  & "SET    newsBookmark = '" & custList(0,I) & "' " _
						  & "WHERE  idNews = " & idNews
				else
					mySQL = "UPDATE newsletters " _
						  & "SET    newsBookmark = NULL " _
						  & "WHERE  idNews = " & idNews
				end if
				set rs = openRSexecute(mySQL)
				call closeRS(rs)
				call closedb()
				
				'Display checkpoint
				Response.Write "Sent -&gt; <b>" & I+1 & "</b> emails <i>(Last email : " & custList(0,I) & ")</i><br>"
				
				'Reset email array counter and values
				I2 = 0
				erase custListEmail
				
				'Is the client still connected?
				if not Response.IsClientConnected then
					Response.End
				end if
				
				'Update Progress Window
				if strUA = "IE" then
					nPctComplete = ( (I+1) / (ubound(custList,2)+1)) * 100
					Response.Write "<SCRIPT LANGUAGE=""JavaScript"">if (progWinOpen==1) progWin.progWinStatus.style.width = " & nPctComplete & ";</SCRIPT>" & vbCrLf
				else
					Response.Write "<SCRIPT LANGUAGE=""JavaScript"">if (progWinOpen==1) progWin.document.frmProgWin.progWinStatus.value = " & I+1 & ";</SCRIPT>" & vbCrLf
				end if
				
				'Send buffered output to browser
				Response.Flush
				
			end if
			
		next
	
		'Close Progress Window
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">if (progWinOpen==1) progWin.close();</SCRIPT>" & vbCrLf
%>
		<br>
		<b>Total number of emails sent : <%=ubound(custList,2)+1%></b><br>
		
	</td></tr>
	</table>
<%
'DISPLAY customer list
else
%>
	<table border=1 cellspacing=1 cellpadding=3>
		<tr>
			<td bgcolor=#EEEEEE><b>Email</b></td>
			<td bgcolor=#EEEEEE><b>Full Name</b></td>
		</tr>
<%
		for I = 0 to ubound(custList,2)
			Response.Write "<tr><td>" & custList(0,I) & "</td><td>" & custList(1,I) & ", " & custList(2,I) & "</td></tr>"
		next
%>
	</table>
<%
end if
%>

<!--#include file="_INCfooter_.asp"-->
