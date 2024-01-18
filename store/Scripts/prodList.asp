<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Displays a list of products that match a given criteria...
'          : - Matches search criteria
'          : - Matches a category
'          : - Matches "specials" on flagged products
'          : If a category is supplied which has sub categories, the
'          : script will display a summary of categories instead of the
'          : product list.
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
dim I
dim totalRecs
dim totalPages
dim count
dim curPage
dim catPos
dim catLst
dim listHeading
dim special
dim strSearch, strSearchMax, strSearchMin, strSearchCat
dim sortField
dim queryStr

'Categories
dim IDCategory
dim categoryDesc
dim IDParentCategory
dim categoryHTML

'Product
dim IDProduct
dim SKU
dim Description
dim DescriptionLong
dim Price
dim Details
dim listPrice
dim smallImageURL
dim imageURL
dim Stock
dim fileName
dim noShipCharge

'Database
dim mySQL
dim conntemp
dim rstemp
dim rstemp2

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

'---------------------------------
' PARMS - Search
'---------------------------------
strSearch    = Request("strSearch")
strSearchMin = Request("strSearchMin")
strSearchMax = Request("strSearchMax")
strSearchCat = Request("strSearchCat")
if len(strSearch & strSearchMin & strSearchMax & strSearchCat) > 0 then

	'Get rid of malicious HTML
	strSearch    = validHTML(strSearch)
	strSearchMin = validHTML(strSearchMin)
	strSearchMax = validHTML(strSearchMax)
	strSearchCat = validHTML(strSearchCat)
	
	'Get rid of multiple spaces in keywords
	do until instr(strSearch,"  ") = 0			
		strSearch = replace(strSearch,"  "," ")
	loop
	
	'If, after all this string manipulation, we have an empty string...
	if len(strSearch & strSearchMin & strSearchMax & strSearchCat) = 0 then
		Response.Clear
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvSearch)
	end if
	
	'Assign default values
	if not(isNumeric(strSearchMin)) then
		strSearchMin = 0
	else
		strSearchMin = CDbl(strSearchMin)
	end if
	if not(isNumeric(strSearchMax)) then
		strSearchMax = 0
	else
		strSearchMax = CDbl(strSearchMax)
	end if
	if not(isNumeric(strSearchCat)) then
		strSearchCat = 0
	else
		strSearchCat = CInt(strSearchCat)
	end if
	
end if

'---------------------------------
' PARMS - Specials
'---------------------------------
special = Request.QueryString("special")
if len(special) > 0 and special <> "Y" then
	special = "N"
end if

'---------------------------------
' PARMS - Categories
'---------------------------------
idCategory = Request.QueryString("idCategory")
if len(idCategory) > 0 then
	'Validate that Category is numeric
	if not IsNumeric(idCategory) then
		Response.Clear
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvCategory)
	end if
	'Validate that Category exists in DB
	mySQL = "SELECT idCategory " _
		  & "FROM   categories " _
		  & "WHERE  idCategory = " & validSQL(idCategory,"I")
	set rsTemp = openRSexecute(mySQL)
	if rsTemp.eof then
		Response.Clear
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvCategory)
	end if
	call closeRS(rsTemp)
end if

'---------------------------------
' PARMS - Validate
'---------------------------------
if  len(strSearch & strSearchMin & strSearchMax & strSearchCat) = 0 _
and len(special) = 0 _
and len(idCategory) = 0 then
	mySQL = "SELECT idCategory " _
		  & "FROM   categories " _
		  & "WHERE  IdParentCategory = 0"
	set rsTemp = openRSexecute(mySQL)
	if rsTemp.eof then
		Response.Clear
		response.redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvCategory & " / " & langErrInvSearch)
	else
		IDCategory = rsTemp("idCategory")
	end if
	call closeRS(rsTemp)
end if

%>
<!--#include file="../UserMods/_INCtop_.asp"--> 
<%

'---------------------------------
' Main driver
'---------------------------------
'SEARCH
if len(strSearch & strSearchMin & strSearchMax & strSearchCat) > 0 then
	listHeading = "<b>" & langGenSearchFor & " : </b> " & strSearch & " [" & strSearchMin & "," & strSearchMax & "," & strSearchCat & "] "
	queryStr	= "strSearch=" & Server.UrlEncode(strSearch) & "&strSearchMin=" & Server.UrlEncode(strSearchMin) & "&strSearchMax=" & Server.UrlEncode(strSearchMax) & "&strSearchCat=" & Server.UrlEncode(strSearchCat)
	call displayItems("search")
else

	'SPECIALS
	if len(special) > 0 then
		listHeading = "<b>" & langGenSpecials & "</b>"
		queryStr	= "special=Y"
		call displayItems("special")
		
	'CATEGORIES
	else
		'Determine category tree position (eg: You are at : cat1 > cat2)
		catPos = getCategoryPos(IDCategory,"")
		
		'Expand the Category tree from the supplied category onward
		catLst = expandCategory(IDCategory,"")
		
		'Display Category Tree position, trim the " <" in catPos
		listHeading = "<b>" & langGenYouAreAt & " : </b>" & mid(catPos,3)
		
		'Display list of products that match category
		if len(trim(catLst)) = 0 then
			queryStr = "idcategory=" & IDCategory
			call displayItems("list")

		'Display Category Tree
		else
			call displayCategory()
		end if
	end if
end if

%>
<!--#include file="../UserMods/_INCbottom_.asp"--> 
<%

call closeDb()

'*************************************************************************
'Determine category position eg: [You are at > cat1 > cat2] (recursive)
'*************************************************************************
function getCategoryPos(IDCategory,tempStr)
	dim mySQL, rsTemp
	mySQL = "SELECT idCategory,idParentcategory,categoryDesc " _
		  & "FROM   categories " _
		  & "WHERE  idCategory = " & validSQL(idCategory,"I")
	set rsTemp = openRSexecute(mySQL)
	do while not rsTemp.eof
		tempStr = " > <a href=""prodList.asp?idCategory=" & rsTemp("idCategory") & """>" & rsTemp("categoryDesc") & "</a>" & tempStr
		tempStr = getCategoryPos(rsTemp("idParentcategory"),tempStr)
		rsTemp.movenext
	loop
	call closeRS(rsTemp)
	getCategoryPos = tempStr
end function
'*************************************************************************
'Expand Categories tree from given category (recursive). Will also 
'display the number of products in each sub category.
'*************************************************************************
function expandCategory(IDCategory,tempStr)

	dim mySQL, rsTemp, catArr, row
	
	'Get Sub-Categories
	mySQL = "SELECT idCategory, categoryDesc,categoryHTML," _
		  & "      (SELECT COUNT(*) " _
		  & "       FROM   products, categories_products " _
		  & "       WHERE  products.idProduct = categories_products.idProduct " _
		  & "       AND    categories_products.idCategory = categories.idCategory " _ 
		  & "       AND    active = -1) " _
		  & "       AS     prodCount " _
		  & "FROM   categories " _
		  & "WHERE  idParentcategory = " & validSQL(idCategory,"I") & " " _
		  & "ORDER BY categoryDesc "
	set rsTemp = openRSexecute(mySQL)
	if not rsTemp.EOF then
	
		'Use getRows() to reduce DB resource requirements. This is a 
		'little more difficult to work with, but makes the queries 
		'much faster. After populating the array, the values are :

		'- catArr(0,row) = idCategory
		'- catArr(1,row) = categoryDesc
		'- catArr(2,row) = categoryHTML
		'- catArr(3,row) = prodCount

		catArr = rsTemp.getRows()
		
	end if
	call closeRS(rsTemp)
	
	'Show Sub-Categories
	if isArray(catArr) then
		tempStr = tempStr & "<ul class=CPcatDescList>"
		for row = 0 to UBound(catArr,2)
			tempStr = tempStr & "<li style=""MARGIN-BOTTOM:5px"">"
			tempStr = tempStr & catArr(2,row)
			if catArr(3,row) = 0 then
				tempStr = tempStr & "<span class=CPcatDesc>" & catArr(1,row) & "</span>"
			else
				tempStr = tempStr & "<span class=CPcatDescProd><a href=""prodList.asp?idCategory=" & catArr(0,row) & """>" & catArr(1,row) & "</a> (" & catArr(3,row) & ")</span>"
			end if
			tempStr = tempStr & "</li>" & vbCrLf
			tempStr = expandCategory(catArr(0,row),tempStr)
		next
		tempStr = tempStr & "</ul>"
	end if	
	
	expandCategory = tempStr
end function
'*************************************************************************
'Display Category Tree
'*************************************************************************
sub displayCategory()
%>
	<table width="100%" border="0" cellspacing="0" cellpadding="2">
		<tr><td valign=middle class="CPpageHead">
			<%=listHeading%><br>
		</td></tr>
	</table>
	
	<img src="../UserMods/misc_cleardot.gif" height=3 width=1><br>

	<table width="100%" border="0" cellspacing="4" cellpadding="4">
		<tr><td valign=top>
			<%=catLst%>
		</td></tr>
	</table>
<%
end sub
'*************************************************************************
'Display list of products for category
'*************************************************************************
sub displayItems(listAction)

	'Determine sort order
	sortField = lcase(trim(Request.QueryString("sortField")))
	if  sortField <> "description" _
	and sortField <> "price" then
		sortField  = "description"
	end if
	
	'Determine page number
	curPage = Request.QueryString("curPage")
	if len(curPage) = 0 or not isNumeric(curPage) then
		curPage = 1
	else
		curPage = CLng(curPage)
	end if

	'Create SQL statement
	select case listAction
	
		'SEARCH
		case "search"
		
			'SQL - General
			mySQL = "SELECT a.idProduct,a.SKU,a.description," _
				  & "       a.descriptionLong,a.listPrice,a.Price," _
				  & "       a.SmallImageUrl,a.Stock,a.fileName," _
				  & "       a.noShipCharge " _
				  & "FROM   products a " _
				  & "WHERE  a.active = -1 "

			'SQL - Minimum Price
			if strSearchMin <> 0 then
				mySQL = mySQL & "AND a.Price >= " & validSQL(strSearchMin,"D") & " "
			end if
			
			'SQL - Maximum Price
			if strSearchMax <> 0 then
				mySQL = mySQL & "AND a.Price <= " & validSQL(strSearchMax,"D") & " "
			end if
			
			'SQL - Category
			if strSearchCat <> 0 then
				mySQL = mySQL _
					  & "AND EXISTS ("_ 
					  & "    SELECT b.idCategory " _
					  & "    FROM   categories_products b " _
					  & "    WHERE  b.idProduct  = a.idProduct " _
					  & "    AND    b.idCategory = " & validSQL(strSearchCat,"I") & ") "
			end if

			'SQL - Keywords
			if len(strSearch) > 0 then
			
				'Declare extra variables
				dim searchArr, tmpSQL1, tmpSQL2, tmpSQL3, tmpSQL4

				'Create array of keywords
				searchArr = split(trim(strSearch)," ")

				'Keyword search SQL
				tmpSQL1 = "(a.details LIKE "
				tmpSQL2 = "(a.description LIKE "
				tmpSQL3 = "(a.descriptionLong LIKE "
				tmpSQL4 = "(a.SKU LIKE "
				for i = 0 to Ubound(searchArr)
					if i = Ubound(searchArr) then
						tmpSQL1 = tmpSQL1 & "'%" & validSQL(searchArr(i),"A") & "%')"
						tmpSQL2 = tmpSQL2 & "'%" & validSQL(searchArr(i),"A") & "%')"
						tmpSQL3 = tmpSQL3 & "'%" & validSQL(searchArr(i),"A") & "%')"
						tmpSQL4 = tmpSQL4 & "'%" & validSQL(searchArr(i),"A") & "%')"
					else
						tmpSQL1 = tmpSQL1 & "'%" & validSQL(searchArr(i),"A") & "%' OR a.details         LIKE "
						tmpSQL2 = tmpSQL2 & "'%" & validSQL(searchArr(i),"A") & "%' OR a.description     LIKE "
						tmpSQL3 = tmpSQL3 & "'%" & validSQL(searchArr(i),"A") & "%' OR a.descriptionLong LIKE "
						tmpSQL4 = tmpSQL4 & "'%" & validSQL(searchArr(i),"A") & "%' OR a.SKU             LIKE "
					end if
				next
				
				'Put it all together
				mySQL = mySQL & "AND (" & tmpSQL1 & " OR " & tmpSQL2 & " OR " & tmpSQL3 & " OR " & tmpSQL4 & ") "
				
			end if
			
			'Sort Order
			mySQL = mySQL & "ORDER BY a." & sortField
			'------------------------------------------------------------
			
		'SPECIALS
		case "special"
		
			mySQL = "SELECT idProduct,SKU,Description,DescriptionLong," _
				  & "       ListPrice,Price,SmallImageUrl,Stock," _
				  & "       fileName,noShipCharge " _
			      & "FROM   products " _
			      & "WHERE  hotDeal = -1 " _
			      & "AND    active = -1 " _
			      & "ORDER BY " & sortField
			
		'CATEGORY
		case else
		
			mySQL = "SELECT a.idProduct,a.SKU,a.Description," _
				  & "       a.DescriptionLong,a.ListPrice,"_
				  & "       a.Price,a.SmallImageUrl,a.Stock," _
				  & "       a.fileName,a.noShipCharge " _
			      & "FROM   products a, categories_products b " _
			      & "WHERE  a.idProduct = b.idProduct " _
			      & "AND    b.idCategory = " & validSQL(idCategory,"I") & " " _
			      & "AND    a.active = -1 " _
			      & "ORDER BY a." & sortField
			
	end select
	
	'Create and Open recordset
	set rsTemp = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pMaxItemsPerPage)

	'Read through recordset and display products
	if rstemp.eof then
		response.write "<br><br><b>" & langErrNoRecFound & "</b><br><br>"
	else
		rstemp.MoveFirst
		rstemp.PageSize		= pMaxItemsPerPage
		totalPages 			= rstemp.PageCount
		totalRecs			= rstemp.RecordCount
		rstemp.AbsolutePage	= curPage
%>
		<table width="100%" border="0" cellspacing="0" cellpadding="2">
			<tr><td valign=middle class="CPpageHead">
				<%=listHeading%><br>
			</td></tr>
		</table>
		<img src="../UserMods/misc_cleardot.gif" height=4 width=1><br>
		
<%
		'Show Page Navigation and Sort if more than one record returned
		if totalRecs > 1 then
%>
			<table width="100%" border="0" cellspacing="0" cellpadding="4" class="CPpageNav">
			<form name="selectPageTopForm">
				<tr>
					<td nowrap align=left valign=middle>
<%
						'Show Page Navigation if more than one page
						if totalPages > 1 then
							call pageNavigation("selectPageTop")
						else
							Response.Write "&nbsp;"
						end if
%>
					</td>
					<td nowrap align=right valign=middle>
						<%call pageSort("sortPageTop")%>
					</td>
				</tr>
			</form>
			</table>
			<img src="../UserMods/misc_cleardot.gif" height=4 width=1><br>
<%
		end if
%>
		<table width="100%" border="0" cellspacing="0" cellpadding="4">
<%
		do while not rstemp.eof and count < rstemp.pageSize
			IDProduct		= rstemp("idProduct")
			SKU             = trim(rstemp("SKU")&"")
			Description		= trim(rstemp("description")&"")
			DescriptionLong	= trim(rstemp("descriptionLong")&"")
			listPrice		= rstemp("listPrice")   
			Price			= rstemp("price")   
			smallImageURL	= trim(rstemp("smallImageUrl")&"")
			Stock      		= rstemp("Stock")
			fileName		= trim(rstemp("fileName")&"")
			noShipCharge	= trim(rstemp("noShipCharge")&"")
%>
			<tr>
			<td align=left valign=top>
				
				<b class="CPprodDesc"><%=Description%></b> 
				<i class="CPprodSKU">(<%=SKU%>)</i><br><br>
				<span class="CPprodDescLong"><%=DescriptionLong%></span><br><br>
					
				<table border=0 cellspacing=0 cellpadding=0>
				<tr>
				<td nowrap valign=bottom width=200>
<%
				'Show Pricing
				if not(pHidePricingZero=-1 and Price=0) then
					if listPrice > Price then
						Response.Write "<span class=CPprodLPriceT>" & langGenListPrice & ":</span> <del class=CPprodLPriceV>" & pCurrencySign & moneyS((listPrice)) & "</del><br>"
					end if
					Response.Write "<b class=CPprodPriceT>" & langGenOurPrice & ":</b> <b class=CPprodPriceV>" & pCurrencySign & moneyS(Price) & "</b>"
					if (listPrice - Price) > 0 then
						Response.Write "<br><span class=CPprodSPriceT>" & langGenYouSave & ":</span> <span class=CPprodSPriceV>" & pCurrencySign & moneyS((listPrice-Price)) & " (" & formatNumber((((listPrice-Price)/listPrice)*100),0) & "%)</span>"
					end if
				end if
%>
				</td>
				<td nowrap valign=bottom>&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td nowrap valign=bottom>
<%
				'Show Extended layout?
				if listViewLayout = 1 then
							
					'Free Shipping message (if not a downloadable item)
					if UCase(noShipCharge) = "Y" and len(fileName) = 0 then
						Response.Write "<b class=CPfreeShipMsg>" & langGenFreeShipping & "</b><br>"
					end if
										
					'In stock, Out of stock message
					if pShowStockView = -1 then
						if pHideAddStockLevel = -1 then
							Response.Write "<b class=CPinStock>" & langGenInStock & "</b><br>"
						else
							if Stock > pHideAddStockLevel then	
								Response.Write "<b class=CPinStock>" & langGenInStock & "</b><br>"
							else
								Response.Write "<b class=CPoutStock>" & langGenOutStock & "</b><br>"
							end if
						end if
					end if
										
					'Show current ratings
					mySQL="SELECT SUM(revRating)   AS revSum,  " _
						& "       COUNT(revRating) AS revCount " _
						& "FROM   reviews " _
						& "WHERE  idProduct = " & validSQL(idProduct,"I") & " " _
						& "AND    revStatus = 'A' "
					set rsTemp2 = openRSexecute(mySQL)
					if not rsTemp2.EOF then
						if rsTemp2("revSum") > 0 and rsTemp2("revCount") > 0 then
							Response.Write "" _
								& "<a href=""prodReview.asp?idProduct=" & idProduct & """>" _
								& langGenAverageRating _
								& "</a> : " _
								& ratingImage(rsTemp2("revSum")/rsTemp2("revCount")) & "<br>"
						end if
					end if
					call closeRS(rsTemp2)
							
				end if
%>
				</td>
				</tr>
				</table>
			</td>
			<td align=center valign=top>
<%
				'Display small product image
				if smallImageURL <> "" then
					Response.Write "<a href=""prodView.asp?idproduct=" & IDProduct & """><img src=""" & pImagesDir & smallImageURL & """ border=0 alt=""" & description & """></a>"
				else
					Response.Write "<b class=CPnoImgT>" & langGenNoImage & "</b>"
				end if
%>
			</td>
			<td nowrap align=center valign=top>
				<a href="prodView.asp?idproduct=<%=IDProduct%>"><img src="../UserMods/butt_view.gif" border="0" align="top" vspace=4></a><br><br>
<%
				'Show View and Add buttons
				if  pCatalogOnly       = 0 and _
				   (pHideAddStockLevel = -1 or _
				    pHideAddStockLevel < CDbl(Stock)) then
						   
					'Check for options and adjust ADD button.
					mySQL = "SELECT idOptionGroup " _
					      & "FROM   optionsGroupsXref " _
					      & "WHERE  idProduct = " & validSQL(idProduct,"I")
					set rsTemp2 = openRSexecute(mySQL)
					if rsTemp2.eof then
%>
						<!-- Use FORM to prevent spiders from adding item to cart -->
						<table border=0 cellpadding=0 cellspacing=0><tr>
						<form method="post" action="cart.asp" name="additem">
							<td>
								<input type="hidden" name="action" value="additem">
								<input type="hidden" name="idProduct" value="<%=IDProduct%>">
								<input type="image"  name="add" src="../UserMods/butt_add.gif" border="0" vspace="4"><br>
							</td>
						</form>
						</tr></table>
<%
					else
%>
						<a href="prodView.asp?idproduct=<%=IDProduct%>"><img src="../UserMods/butt_add.gif" border="0" align="top" vspace=4></a><br>
<%
					end if
%>
					<br><img border="0" height="1" width="60" src="../UserMods/misc_cleardot.gif"><br>
<%
					call closeRS(rsTemp2)
				end if
%>
			</td>
			</tr>
<%   
			count = count + 1  
			rstemp.moveNext
			
			'Draw line between products
			if not rstemp.EOF and count < rstemp.pageSize then
%>
			<tr>
				<td colspan=3>
					<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr><td class="CPlines" width="100%">
							<img border="0" height="1" width="1" src="../UserMods/misc_cleardot.gif"><br>
						</td></tr>
					</table>
				</td>
			</tr>
<%
			end if
			
		loop
%>
		</table>
		<img src="../UserMods/misc_cleardot.gif" height=4 width=1><br>

<%
		'Show Page Navigation if more than one page returned
		if totalPages > 1 then
%>
			<table width="100%" border="0" cellspacing="0" cellpadding="4" class="CPpageNav">
			<form name="selectPageBotForm">
				<tr><td valign=middle>
					<%call pageNavigation("selectPageBot")%>
				</td></tr>
			</form>
			</table>
<%
		else
%>
			<table border=0 cellpadding=0 cellspacing=0 width="100%">
				<tr><td class="CPlines" width="100%">
					<img border="0" height="1" width="1" src="../UserMods/misc_cleardot.gif"><br>
				</td></tr>
			</table>
<%
		end if
	end if
	call closeRS(rsTemp)
end sub
'*********************************************************************
'Display page navigation
'*********************************************************************
sub pageNavigation(formFieldName)
	Response.Write langGenNavPage & " "
	Response.Write "<select onChange=""location.href=this.options[selectedIndex].value"" name=" & trim(formFieldName) & " style=""FONT-SIZE: 8pt"">"
	for I = 1 to TotalPages
		Response.Write "<option value=""prodList.asp?" & queryStr & "&curPage=" & I & "&sortField=" & server.URLEncode(sortField) & """ " & checkMatch(curPage,I) & ">" & I & "</option>" & vbCrlf
	next
	Response.Write "</select> " & langGenOf & " " & TotalPages & "&nbsp;&nbsp;"
	Response.Write "[ "
	if curPage > 1 then
		Response.Write "<a href=""prodList.asp?" & queryStr & "&curPage=" & curPage-1 & "&sortField=" & server.URLEncode(sortField) & """>" & langGenNavBack & "</a>"
	else
		Response.Write langGenNavBack
	end if
	Response.Write " | "
	if curPage < TotalPages then
		Response.Write "<a href=""prodList.asp?" & queryStr & "&curPage=" & curPage+1 & "&sortField=" & server.URLEncode(sortField) & """>" & langGenNavNext & "</a>"
	else
		Response.Write langGenNavNext
	end if
	Response.Write " ]"
end sub
'*********************************************************************
'Display sort list
'*********************************************************************
sub pageSort(formFieldName)
	Response.Write langGenSort & " : "
%>
	<select onChange="location.href=this.options[selectedIndex].value" name="<%=trim(formFieldName)%>" style="FONT-SIZE: 8pt">
		<option value="prodList.asp?<%=queryStr & "&sortField=description"%>" <%=checkMatch(sortField,"description")%>><%=langGenItemDesc%></option>
		<option value="prodList.asp?<%=queryStr & "&sortField=price"	  %>" <%=checkMatch(sortField,"price")%>><%=langGenPrice%></option>
	</select>
<%
end sub
%>
