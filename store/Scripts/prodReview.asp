<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Displays and captures reviews for a product
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
%>
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="_INCconfig_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="_INCappFunctions_.asp"-->  
<%
'Work Fields
dim totalPages
dim count
dim curPage
dim reviewsPerPage
dim revCount
dim revSum
dim I

'Reviews
dim idReview
dim revDate
dim revDateInt
dim revAuditInfo
dim revStatus
dim revRating
dim revName
dim revLocation
dim revEmail
dim revSubj
dim revDetail

'Product
dim idProduct
dim sku
dim description
dim reviewAutoActive

'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2

'Session
dim idOrder
dim idCust

'Set number of reviews per page
reviewsPerPage = 10

'*************************************************************************

'Open Database Connection
call openDb()

'Store Configuration
if loadConfig() = false then
	call errorDB(langErrConfig,"")
end if

'Get/Set Cart/Order Session
idOrder = sessionCart()

'Get/Set Customer Session
idCust  = sessionCust()

'Get idProduct and validate
idProduct = Request.QueryString("idProduct")
if len(idProduct) = 0 then
	idProduct = Request.Form("idProduct")
end if
if IsNumeric(idProduct) then
	mySQL = "SELECT sku,description,reviewAutoActive " _
	      & "FROM   products " _
	      & "WHERE  idProduct = " & validSQL(idProduct,"I") & " " _
	      & "AND    active = -1 " _
	      & "AND    reviewAllow = 'Y' "
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.eof then
		sku              = rstemp("sku")
		description	     = rstemp("description")
		reviewAutoActive = rstemp("reviewAutoActive")
	else
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvProdID)
	end if
	call closeRS(rsTemp)
else
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvProdID)
end if

'Did the customer add a new review?
if Request.Form("formID") <> "" then
	call newReviewAdd()
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<table width="100%" border="0" cellspacing="0" cellpadding="2">
	<tr><td valign=middle class="CPpageHead">
		<b><%=langGenProductReviews%></b>
	</td></tr>
</table>

<!-- Product Description and Review Summary -->

<table width="100%" border="0" cellspacing="4" cellpadding="4">
	<tr><td>
		<b class="CPprodDesc"><a href="prodView.asp?idProduct=<%=idProduct%>"><%=SKU%></a> - <%=Description%></b>
	</td></tr>
<%
	'Get Ratings Summary
	mySQL="SELECT SUM(revRating)   AS revSum,  " _
		& "       COUNT(revRating) AS revCount " _
		& "FROM   reviews " _
		& "WHERE  idProduct = " & validSQL(idProduct,"I") & " " _
		& "AND    revStatus = 'A' "
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.EOF then
		revSum   = rsTemp("revSum")
		revCount = rsTemp("revCount")
		if revSum > 0 and revCount > 0 then
%>
			<tr><td>
				<b><%=langGenAverageRating%> :</b> <%=ratingImage(revSum/revCount)%><br>
				<b><%=langGenNumberReviews%> : <%=revCount%></b><br>
				<b><a href="#write"><%=langGenWriteReview%></a></b><br>
			</td></tr>
<%
		end if
	end if
	call closeRS(rsTemp)
%>
</table>

<!-- Product Review Detail -->

<table width="100%" border="0" cellspacing="4" cellpadding="4">
<%	
	'Get Reviews
	mySQL = "SELECT revDate,revRating,revName,revLocation," _
		  & "       revEmail,revSubj,revDetail " _
	      & "FROM   reviews " _
	      & "WHERE  idProduct = " & validSQL(idProduct,"I") & " " _
	      & "AND    revStatus = 'A' " _
	      & "ORDER BY revDate DESC"
	
	'Create and Open recordset
	set rsTemp = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,reviewsPerPage)

	'Check if any records were returned
	if not rstemp.eof then
	
		'Get Page to show (if any)
		curPage = Request.Form("curPage")
		if len(curPage) = 0 then
			curPage = Request.QueryString("curPage")
		end if
		if len(curPage) = 0 or not isNumeric(curPage) then
			curPage = 1
		else
			curPage = CLng(curPage)
		end if
	
		'Go to requested page
		rstemp.MoveFirst
		rstemp.PageSize		= reviewsPerPage
		totalPages 			= rstemp.PageCount
		rstemp.AbsolutePage	= curPage
%>
		<tr><td class="CPpageNav">
			<%=navbarReviews("prodReview.asp","idProduct=" & idProduct)%>
		</td></tr>
<%
		'Read through recordset and display reviews
		do while not rstemp.eof and count < rstemp.pageSize
			revDetail	= rstemp("revDetail")
			revDate     = rstemp("revDate")
			revRating   = rstemp("revRating")
			revName		= rstemp("revName")
			revLocation	= rstemp("revLocation")
			revEmail	= rstemp("revEmail")
			revSubj		= rstemp("revSubj")
%>
			<tr>
				<td valign=top>
					<%=ratingImage(revRating)%>&nbsp;
					<b><%=server.HTMLEncode(revSubj)%></b><br>
					<i><%=server.HTMLEncode(revName)%> - <%=server.HTMLEncode(revLocation)%>&nbsp;&nbsp; (<%=formatDateTime(revDate,vbLongDate)%>)</i><br><br>
					<%=replace(server.HTMLEncode(revDetail),chr(10),"<br>")%><br><br>
				</td>
			</tr>
<%   
			count = count + 1  
			rstemp.moveNext
		loop
	else
%>
		<tr><td>
			<%=langGenNotReviewedYet%>
		</td></tr>
<%
	end if
	call closeRS(rsTemp)
%>
</table>

<!-- New Review Form -->

<a name="write">

<table width="100%" border="0" cellspacing="0" cellpadding="2">
	<tr><td valign=middle class="CPpageHead">
		<b><%=langGenWriteReview%></b>
	</td></tr>
</table>

<img src="../UserMods/misc_cleardot.gif" height=4 width=1><br>

<table width="100%" border="0" cellspacing="0" cellpadding="2">
<form method="post" name="prodReview" action="prodReview.asp">
	<tr>
		<td>Rating</td>
		<td valign=middle nowrap>
			<select name=revRating id=revRating size=1>
				<option value=""></option>
				<option value="5">5</option>
				<option value="4">4</option>
				<option value="3">3</option>
				<option value="2">2</option>
				<option value="1">1</option>
			</select> 
			<%=langGenReviewStars%>
		</td>
	</tr>
	<tr>
		<td><%=langGenName%></td>
		<td><input type="text" name="revName" size="20" maxlength="250"></td>
	</tr>
	<tr>
		<td><%=langGenLocation%></td>
		<td><input type="text" name="revLocation" size="20" maxlength="250"></td>
	</tr>
	<tr>
		<td><%=langGenEmail%></td>
		<td><input type="text" name="revEmail" size="20" maxlength="100"></td>
	</tr>
	<tr>
		<td><%=langGenSubject%></td>
		<td><input type="text" name="revSubj" size="20" maxlength="100"></td>
	</tr>
	<tr>
		<td><%=langGenReview%></td>
		<td><textarea name="revDetail" rows=6 cols="40" wrap="soft"></textarea></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<input type="hidden" name="idProduct" value="<%=idProduct%>">
			<input type="hidden" name="formID"    value="00">
			<input type="submit" name="submit"    value="<%=langGenSend%>">
		</td>
	</tr>
</form>
</table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"--> 

<%

call closeDb()

'**********************************************************************
'Add new review to database
'**********************************************************************
sub newReviewAdd()

	'Check if this customer has reviewed this product before
	mySQL="SELECT revAuditInfo " _
		& "FROM   reviews " _
		& "WHERE  idProduct    = "  & validSQL(idProduct,"I") & " " _
		& "AND    revAuditInfo = '" & validSQL(Request.ServerVariables("REMOTE_ADDR"),"A") & "' "
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.EOF then
		'Say thank you, even though we completely ignore the review. 
		'This is a common practice and is part of the effort to 'hide' 
		'this software's anti-spam mechanisms from the spammer.
		response.redirect "sysMsg.asp?msg=" & server.URLEncode(langGenReviewAddedMsg) & "&returnURL=" & server.URLEncode("prodReview.asp?idProduct=" & idProduct)
	end if
	call closeRS(rsTemp)

	'Get form values
	revRating   = validHTML(request.Form("revRating"))
	revName     = validHTML(request.Form("revName"))
	revLocation = validHTML(request.Form("revLocation"))
	revEmail    = validHTML(request.Form("revEmail"))
	revSubj     = validHTML(request.Form("revSubj"))
	revDetail   = validHTML(request.Form("revDetail"))
	
	'Check form values
	if len(revRating)  = 0 _
	or len(revName)    = 0 _
	or len(revLocation)= 0 _
	or len(revEmail)   = 0 _
	or len(revSubj)    = 0 _
	or len(revDetail)  = 0 _
	or invalidChar(revRating,0,"12345") _
	or invalidChar(revEmail,1,"@.-_") then
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvForm)
	end if
	
	'Check if review must be activated automatically
	if reviewAutoActive = "Y" then
		revStatus = "A"
	else
		revStatus = "I"
	end if
	
	'INSERT review record
	mySQL = "INSERT INTO reviews (" _
		  & "idProduct,revDate,revDateInt,revAuditInfo,revStatus," _
		  & "revRating,revName,revLocation,revEmail,revSubj,revDetail) " _
		  & "VALUES (" _
	      &       validSQL(idProduct,"I")		& ", " _
	      &	"'" & validSQL(now(),"A")			& "'," _
	      &	"'" & validSQL(dateInt(now()),"A")	& "'," _
	      &	"'" & validSQL(Request.ServerVariables("REMOTE_ADDR"),"A") & "'," _
	      &	"'" & validSQL(revStatus,"A")		& "'," _
	      &       validSQL(revRating,"I")		& ", " _
	      &	"'" & validSQL(revName,"A")			& "'," _
	      &	"'" & validSQL(revLocation,"A")		& "'," _
	      &	"'" & validSQL(revEmail,"A")		& "'," _
	      &	"'" & validSQL(revSubj,"A")			& "'," _
	      &	"'" & validSQL(revDetail,"A")		& "' " _
	      & ")"
	set rsTemp = openRSexecute(mySQL)
	call closeRS(rsTemp)
	
	'Say thank you
	response.redirect "sysMsg.asp?msg=" & server.URLEncode(langGenReviewAddedMsg) & "&returnURL=" & server.URLEncode("prodReview.asp?idProduct=" & idProduct)
	
end sub
'**********************************************************************
'Display navigation bar
'**********************************************************************
function navbarReviews(scriptName,queryParms)

	'Page number
	Response.Write langGenNavPage & " : " & curPage & " / " & TotalPages & " &nbsp;&nbsp; "
		
	'Back Button
	if curPage > 1 then
		Response.Write "[ <a href=""" & scriptName & "?" & queryParms & "&curPage=" & curPage-1 & """>" & langGenNavBack & "</a>"
	else
		Response.Write "[ " & langGenNavBack
	end if
		
	'Next Button
	if curPage < TotalPages then
		Response.Write " | <a href=""" & scriptName & "?" & queryParms & "&curPage=" & curPage+1 & """>" & langGenNavNext & "</a>" & " ]"
	else
		Response.Write " | " & langGenNavNext & " ]"
	end if

end function
%>
