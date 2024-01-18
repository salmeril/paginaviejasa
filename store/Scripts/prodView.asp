<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Display details for a specific product, including all 
'          : options.
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
'Product
dim IDProduct
dim Description
dim DescriptionLong
dim Price
dim Details
dim relatedKeys
dim listPrice
dim smallImageURL
dim imageURL
dim Stock
dim SKU
dim fileName
dim noShipCharge
dim reviewAllow

'Options
dim priceToAdd
dim percToAdd

'Work Fields
dim testMode
dim revCount
dim revSum

'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2
dim rsTemp3

'Session
dim idOrder
dim idCust

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

'Check Product Code
IDProduct = request.QueryString("idProduct")
if trim(IDProduct) = "" or not IsNumeric(IDProduct) then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvProdID)
end if

'Check if Product is being displayed as a Test
if Instr(lCase(Request.ServerVariables("HTTP_REFERER")),"/sa_prod") = 0 then
	testMode = false
else
	testMode = true
end if

'Get Product Detail
mySQL = "SELECT description,descriptionLong,relatedKeys,price," _
	  & "       listprice,smallImageUrl,imageurl,stock,sku," _
	  & "       fileName,noShipCharge,reviewAllow,details " _
	  & "FROM   products " _
	  & "WHERE  idProduct = " & validSQL(idProduct,"I") & " "
if not testMode then
	mySQL = mySQL & "AND active = -1"
end if
set rsTemp = openRSexecute(mySQL)
if rsTemp.eof then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvProdID)
end if

'Assign product DB fields to local fields
Details		    = trim(rsTemp("details")&"")
Description	    = trim(rsTemp("description")&"")
DescriptionLong = trim(rsTemp("DescriptionLong")&"")
relatedKeys     = trim(rsTemp("relatedKeys")&"")
Price			= rsTemp("price")
listPrice		= rsTemp("listPrice")
smallImageURL	= trim(rsTemp("SmallImageUrl")&"")
imageURL		= trim(rsTemp("imageUrl")&"")
Stock			= rsTemp("stock")
sku			    = trim(rsTemp("sku")&"")
fileName		= trim(rsTemp("fileName")&"")
noShipCharge    = trim(rsTemp("noShipCharge")&"")
reviewAllow	    = trim(rsTemp("reviewAllow")&"")

call closeRS(rsTemp)

%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<table border="0" cellpadding=0 cellspacing=0 width="100%">
	<form method="post" action="cart.asp" name="additem">
	<tr>
<%
		if prodViewLayout = "0" then
%>
		<td align=left valign=top width="64%">
			<b class="CPprodDescDet"><%=Description%></b> 
			<i class="CPprodSKU">(<%=SKU%>)</i><br>
			<%call getProdDetail()%>
			<%call getProdPricing()%>
			<%call getProdStock()%>
			<%call getFreeShipMsg()%>
			<%call getOptionsGroups()%>
			<%call getQtyAndAdd()%>
			<%call getProdDisc()%>
			<%call getProdRelated()%>
			<%call getProdReview()%>
		</td>
		<td width="2%">&nbsp;&nbsp;</td>
		<td align=left valign=top width="34%">
			<%call getProdImage()%>
			<%call getEmailToFriend()%>
			<%call getContactUs()%>
		</td>
<%
		else
%>
		<td align=left valign=top width="49%">
			<b class="CPprodDescDet"><%=Description%></b> 
			<i class="CPprodSKU">(<%=SKU%>)</i><br>
			<%call getProdImage()%>
			<%call getProdDetail()%>
			<%call getProdRelated()%>
			<%call getProdReview()%>
		</td>
		<td width="2%">&nbsp;&nbsp;</td>
		<td align=left valign=top width="49%">
			<table border=0 cellSpacing=0 cellPadding=5 width=180 class="CPbox2">
				<tr>
					<td nowrap align=center class="CPbox2H">
						<b><%=langGenReadyToOrder%></b>
					</td>
				</tr>
				
				<tr>
					<td nowrap align=left valign=top class="CPbox2B">
						<%call getProdPricing()%>
						<%call getProdStock()%>
						<%call getFreeShipMsg()%>
						<%call getOptionsGroups()%>
						<%call getQtyAndAdd()%>
						<%call getProdDisc()%>
					</td>
				 </tr>
				 <tr>
					<td nowrap align=center class="CPbox2H">
						<b>SDFASD</b>
					</td>
				</tr>
			</table>
			<%call getEmailToFriend()%>
			<%call getContactUs()%>
		</td>
<%
		end if
%>
	</tr>
	</form>  
</table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"--> 
<%

call closeDb()

'*************************************************************************
' Display Product Details
'*************************************************************************
sub getProdDetail()
	if len(Details) = 0 then
%>
		<br><span class=CPprodDet><%=DescriptionLong%></span><br>
<%
	else
%>
		<br><span class=CPprodDet><%=Details%></span><br>
<%
	end if
end sub
'*************************************************************************
' Display Product Pricing
'*************************************************************************
sub getProdPricing()

	'Check if we need to hide pricing if price=0
	if pHidePricingZero=-1 and Price=0 then
		exit sub
	end if
	
	'Show Prices
	if listPrice > Price then
%>
		<br><span class=CPprodLPriceT><%=langGenListPrice%>:</span> <del class=CPprodLPriceV><%=pCurrencySign & moneyS((listPrice))%></del>
<%
	end if
%>
	<br><b class=CPprodPriceT><%=langGenOurPrice%>:</b> <b class=CPprodPriceV><%=pCurrencySign & moneyS(Price)%></b>
<%
	if (listPrice - Price) > 0 then
%>
		<br><span class=CPprodSPriceT><%=langGenYouSave%>:</span> <span class=CPprodSPriceV><%=pCurrencySign & moneyS((listPrice-Price))%> (<%=formatNumber((((listPrice-Price)/listPrice)*100),0)%>%)</span>
<%
	end if
%>
	<br>
<%
end sub
'*************************************************************************
' Display Product Stock Description
'*************************************************************************
sub getProdStock()
	if pShowStockView = -1 then
		if pHideAddStockLevel = -1 then
%>
			<br><b class=CPinStock><%=langGenInStock%></b><br>
<%
		else
			if Stock > pHideAddStockLevel then
%>
				<br><b class=CPinStock><%=langGenInStock%></b><br>
<%
			else
%>
				<br><b class=CPoutStock><%=langGenOutStock%></b><br>
<%
			end if
		end if
	end if
end sub
'*************************************************************************
' Display Free Shipping message
'*************************************************************************
sub getFreeShipMsg()
	'Check if "Free Shipping" and NOT a downloadable item
	if UCase(noShipCharge) = "Y" and len(fileName) = 0 then
%>
		<br><b class=CPfreeShipMsg><%=langGenFreeShipping%></b><br>
<%
	end if
end sub
'*************************************************************************
' Display Product Options
'*************************************************************************
sub getOptionsGroups()

	dim mySQL, rstemp, rstemp2, rstemp3, optionDesc

	'Get option groups for this Product
	mySQL = "SELECT a.idOptionGroup, a.optionGroupDesc, " _
		  & "       a.optionReq, a.optionType " _
	      & "FROM   optionsGroups a, optionsGroupsXref b " _
	      & "WHERE  a.idOptionGroup = b.idOptionGroup " _
	      & "AND    b.idProduct = " & validSQL(idProduct,"I") & " " _
	      & "ORDER BY a.optionGroupDesc "
	set rsTemp2 = openRSexecute(mySQL)
	
	'Extra line break before displaying option groups
	if not rsTemp2.EOF then
		Response.Write "<br>" & vbCrlf
	end if
	
	'Loop through option groups
	do while not rstemp2.EOF
		
		'Get Options for Option Group
		mySQL = "SELECT b.idOption, b.optionDescrip, b.priceToAdd, " _
			  & "       b.percToAdd " _
		      & "FROM   optionsXref a, options b " _
		      & "WHERE  a.idOptionGroup = " & rstemp2("idOptionGroup") & " " _
		      & "AND    b.idOption = a.idOption " _
		      & "AND    NOT EXISTS " _
		      & "      (SELECT c.idOptionsProdEx " _
		      & "       FROM   OptionsProdEx c " _
		      & "       WHERE  c.idOption = b.idOption " _
		      & "       AND    c.idProduct = " & validSQL(idProduct,"I") & ") "
		set rsTemp3 = openRSexecute(mySQL)
 		if not rstemp3.EOF then
	
			'Display option group description and create hidden form 
			'variables to be used in validations and error messages.
			if UCase(rstemp2("optionReq")) = "Y" then
				Response.Write "<span class=CPoptDesc>" & rstemp2("optionGroupDesc") & " :</span><br>" & vbCrlf
			else
				Response.Write "<span class=CPoptDesc>" & rstemp2("optionGroupDesc") & " (" & langGenOptional & ") :</span><br>" & vbCrlf
			end if
			Response.Write "<input type=""hidden"" name=""DESidOption" & rstemp2("idOptionGroup") & """ value=""" & rstemp2("optionGroupDesc") & """>" & vbCrlf
			Response.Write "<input type=""hidden"" name=""REQidOption" & rstemp2("idOptionGroup") & """ value=""" & rstemp2("optionReq")       & """>" & vbCrlf
			Response.Write "<input type=""hidden"" name=""TYPidOption" & rstemp2("idOptionGroup") & """ value=""" & rstemp2("optionType")      & """>" & vbCrlf
 		
 			'Drop-down list of options
 			if rstemp2("optionType") = "S" then
				Response.Write "<select name=""OPTidOption" & rstemp2("idOptionGroup") & """ class=""CPoptSel"">" & vbCrlf
				Response.Write "<option value=""""></option>" & vbCrlf
				do while not rstemp3.EOF
					priceToAdd = getOptionPrice(rstemp3("priceToAdd"),rstemp3("percToAdd"),price)
					if priceToAdd > 0 then
						optionDesc = rstemp3("optionDescrip") & " " & pCurrencySign & moneyS(priceToAdd)
					else
						optionDesc = rstemp3("optionDescrip")
					end if
					if   optionDesc >= "as" then
					Response.Write "<option value=""" & rstemp3("idOption") & """selected>" & optionDesc & "</option>" & vbCrlf
					else
					Response.Write "<option value=""" & rstemp3("idOption") & """>" & optionDesc & "</option>" & vbCrlf
					end if
					rstemp3.movenext
				loop
				Response.Write "</select><br>" & vbCrlf
			end if
				
			'Text input option
 			if rstemp2("optionType") = "T" then
				Response.Write "<input type=""hidden"" name=""OPTidOption" & rstemp2("idOptionGroup") & """ value=""" & rstemp3("idOption") & """>" & vbCrlf
				Response.Write "<input type=""text""   name=""TXTidOption" & rstemp2("idOptionGroup") & """ size=""25"" maxlength=""200"" class=""CPoptTxt""><br>" & vbCrlf
			end if
			
		end if
		call closeRS(rsTemp3)
		rstemp2.movenext
	loop
	call closeRS(rsTemp2)

end sub
'*************************************************************************
' Display Qty box and Add Button
'*************************************************************************
sub getQtyAndAdd()
	if pCatalogOnly = 0 and _
	  (pHideAddStockLevel = -1 or _
	   pHideAddStockLevel < CDbl(Stock)) then
%>
		<br>
		<input type="hidden" name="action"     value="additem">
		<input type="hidden" name="idProduct"  value="<%=IDProduct%>">
		<input type="text"   name="quantity"   value="1" size="2" maxlength="2"> &nbsp;
		<input type="image"  name="add"        src="../UserMods/butt_add.gif" border="0">
		<br>
<%
	end if
end sub
'*************************************************************************
' Display Product Discounts
'*************************************************************************
sub getProdDisc()

	dim mySQL, rstemp

	'Get Product Discounts
	mySQL="SELECT discAmt,discFromQty,discToQty,discPerc " _ 
		& "FROM   DiscProd " _
		& "WHERE  idProduct = " & validSQL(idProduct,"I") & " " _
		& "ORDER BY discFromQty"
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.EOF then
%>
		<br>
		<table border=0 cellpadding=0 cellspacing=0>
			<tr><td>
				<b><%=langGenDiscount%> (<%=langGenQty%>) :</b>
			</td></tr>
			<tr><td>
				<table border=0 cellpadding=0 cellspacing=0>
<%
				do while not rsTemp.EOF
					Response.Write "" _
						& "<tr>" _
						& "  <td nowrap>" & rsTemp("discFromQty") & "</td>" _
						& "  <td nowrap>&nbsp;-&nbsp;</td>" _
						& "  <td nowrap>" & rsTemp("discToQty") & "</td>" _
						& "  <td nowrap>&nbsp&nbsp;</td>" _
						& "  <td nowrap>"
					if isNull(rsTemp("discPerc")) then
						Response.Write pCurrencySign & moneyS(rsTemp("discAmt")) & " ea."
					else
						Response.Write rsTemp("discPerc") & "% ea."
					end if
					Response.Write "" _
						& "  </td>" _
						& "</tr>"
					rsTemp.Movenext 
				loop
%>
				</table>
			</td></tr>
		</table>
<%
	end if
	call closeRS(rsTemp)
end sub
'*************************************************************************
' Display Related Products
'*************************************************************************
sub getProdRelated()

	dim mySQL, rstemp, rsTemp2
%>
	<br>
	<table border=0 cellpadding=0 cellspacing=0>
		<tr><td>
			<b><%=langGenRelatedProd%> :</b>
		</td></tr>
		<tr><td>
<%
		'Get other Product Group products
		mySQL="SELECT prodGroupP " _
			& "FROM   productGroups " _
			& "WHERE  prodGroupC  = " & validSQL(idProduct,"I")
		set rsTemp = openRSexecute(mySQL)
		do while not rsTemp.EOF
			mySQL="SELECT b.idProduct, b.description," _
				& "       b.price " _
				& "FROM   productGroups a, products b " _
				& "WHERE  a.prodGroupP = " & rsTemp("prodGroupP") & " " _
				& "AND    a.prodGroupC <> " & validSQL(idProduct,"I") & " " _
				& "AND    b.idProduct  = a.prodGroupC " _
				& "AND    b.active = -1 " _
				& "ORDER BY b.description "
			set rsTemp2 = openRSexecute(mySQL)
			do while not rstemp2.EOF
				Response.Write "<a href=""prodView.asp?idProduct=" & rsTemp2("idProduct") & """>" & rsTemp2("description") & "</a> - " & pCurrencySign & moneyS(rsTemp2("Price")) & "<br>"
				rsTemp2.Movenext
			loop
			call closeRS(rsTemp2)
			rsTemp.Movenext
		loop
		call closeRS(rsTemp)

		'Get categories for this product
		mySQL="SELECT a.idCategory, b.categoryDesc " _
			& "FROM   Categories_Products a " _
			& "INNER JOIN Categories b " _
			& "ON     a.idCategory = b.idCategory " _
			& "WHERE  a.idProduct = " & validSQL(idProduct,"I")
		set rsTemp = openRSexecute(mySQL)
		if not rsTemp.EOF then
			do while not rsTemp.eof
				Response.Write "<a href=""prodList.asp?idCategory=" & rsTemp("idCategory") & """>" & langGenCategory & " : " & rsTemp("categoryDesc") & "</a><br>"
				rsTemp.movenext
			loop
		else
			Response.Write langGenCategory & " : " & langGenNotApplicable & "<br>"
		end if
		call closeRS(rsTemp)
				
		'Related keys for this product
		if len(relatedKeys) > 0 then
			Response.Write "<a href=""prodList.asp?strSearch=" & server.URLEncode(relatedKeys) & """>" & langGenSearchRelated & "</a><br>"
		end if
%>
		</td></tr>
	</table>
<%
end sub
'*************************************************************************
' Display Product Review Summary
'*************************************************************************
sub getProdReview()

	dim mySQL, rstemp
	
	'Check if reviews are allowed for this product
	if UCase(reviewAllow) = "Y" then
%>
		<br>
		<table border=0 cellpadding=0 cellspacing=0>
			<tr><td>
				<b><%=langGenProductReviews%> :</b>
			</td></tr>
			<tr><td>
<%			
			'Get current ratings
			mySQL="SELECT SUM(revRating)   AS revSum,  " _
				& "       COUNT(revRating) AS revCount " _
				& "FROM   reviews " _
				& "WHERE  idProduct = " & validSQL(idProduct,"I") & " " _
				& "AND    revStatus = 'A' "
			set rsTemp = openRSexecute(mySQL)
			if not rsTemp.EOF then
				revSum   = rsTemp("revSum")
				revCount = rsTemp("revCount")
			else
				revSum   = 0
				revCount = 0
			end if
			call closeRS(rsTemp)
				
			'Show Ratings
			if revSum > 0 and revCount > 0 then
				Response.Write "" _
				& langGenAverageRating & " : " _
				& ratingImage(revSum/revCount) & "<br>" _
				& "<a href=""prodReview.asp?idProduct=" & idProduct & """>" & langGenNumberReviews & "</a> : " & revCount & "<br>"
			end if
					
			'Rate this link
			Response.Write "<a href=""prodReview.asp?idProduct=" & idProduct & """>" & langGenWriteReview & "</a><br>"
%>
			</td></tr>
		</table>
<%
	end if
end sub
'*************************************************************************
' Display Product Image
'*************************************************************************
sub getProdImage()
	if imageURL <> "" then
%>
		<br><img src="<%=pImagesDir & imageURL%>" border=0 alt="<%=description%>"><br>
<%
	else
		if smallImageURL <> "" then
%>
			<br><img src="<%=pImagesDir & smallImageURL%>" border=0 alt="<%=description%>"><br>
<%
		else
%>
			<br><b class=CPnoImgT><%=langGenNoImage%></b><br>
<%
		end if
	end if
end sub
'*************************************************************************
' Display "Email To Friend" link
'*************************************************************************
sub getEmailtoFriend()
	if mailComp <> 0 or demoMode = "Y" then
%>
		<br><img src="../UserMods/misc_Email.gif" border="0"> <a href="emailToFriend.asp?idProduct=<%=server.URLEncode(IDProduct)%>&description=<%=server.URLEncode(Description)%>&price=<%=server.URLEncode(Price)%>"><%=langGenEmailFriendHdr%></a><br>
<%
	end if
end sub
'*************************************************************************
' Display "Contact Us" link
'*************************************************************************
sub getContactUs()
	if mailComp <> 0 or demoMode = "Y" then
%>
		<br><img src="../UserMods/misc_Email.gif" border="0"> <a href="contactUs.asp?emailSubject=<%=server.URLEncode(langGenProdInquiry & " : " & SKU & " - " & Description)%>"><%=langGenProdInquiry%></a><br>
<%
	end if
end sub
%>
