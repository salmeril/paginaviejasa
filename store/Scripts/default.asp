<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : "Home" Page for Store Front
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
call openDB()

'Store Configuration
if loadConfig() = false then
	call errorDB(langErrConfig,"")
end if

'Get/Set Cart/Order Session
idOrder = sessionCart()

'Get/Set Customer Session
idCust  = sessionCust()

%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<table border=0 cellspacing=0 cellpadding=0 width="100%">
	<tr>
		<td align=left valign=top>
			<!--#include file="../UserMods/_INCleft_.asp"-->
		</td>
		<td align=left valign=top width="100%">
			<table border=0 cellpadding=5 cellspacing=0 width="100%">
<%
			'Get Fetured Products
			mySQL="SELECT idProduct,description,descriptionLong," _
			    & "       listPrice,price,smallImageUrl,stock," _
			    & "       fileName,noShipCharge " _
			    & "FROM   products " _
			    & "WHERE  active = -1 " _
			    & "AND    homePage = -1 " _
			    & "ORDER BY idProduct DESC "
			set rsTemp = openRSopen(mySQL,0,adOpenStatic,adLockReadOnly,adCmdText,pMaxItemsPerPage)
			
			'If Featured Products = 0 then use "_INCright_.asp"
			if rsTemp.EOF then
				call noProd()
			else
			
				'If Featured Products <= 2 then display single rows
				if rsTemp.RecordCount <= 2 then
					do while not rsTemp.EOF
						call singleProd()
						if not rsTemp.EOF then
							rsTemp.MoveNext
						end if
					loop
					
				'If Featured Products > 2 then display double rows
				else
					do while not rsTemp.EOF
						call doubleProd()
						if not rsTemp.EOF then
							rsTemp.MoveNext
						end if
					loop
					
				end if					
			end if
%>
			</table>
		</td>
	</tr>
</table>

<!--#include file="../UserMods/_INCbottom_.asp"-->
<%
call closeDB()

'**********************************************************************
'Displayed when there are no Featured Products.
'**********************************************************************
sub noProd()
%>
	<tr>
		<td align=center valign=top>
			<!--#include file="../UserMods/_INCright_.asp"-->
		</td>
	</tr>
<%
end sub
'**********************************************************************
'Displays ONE Featured Product across the width of a row.
'**********************************************************************
sub singleProd()
%>
	<tr>
		<td nowrap align=center valign=top class="CPhomeImg">
			<%call prodImage()%>
		</td>
		<td align=left valign=top class="CPhomeDesc">
			<%call prodDetail()%>
		</td>
	</tr>
	<tr><td colspan=2><img src="../UserMods/misc_cleardot.gif" height=1 width=1></td></tr>
<%
end sub
'**********************************************************************
'Displays TWO Featured Products across the width of a row.
'**********************************************************************
sub doubleProd()
%>
	<tr>
		<td nowrap valign=top align=center class="CPhomeImg">
			<%call prodImage()%>
		</td>
		<td valign=top class="CPhomeDesc" style="width:50%">
			<%call prodDetail()%>
		</td>
		<td style="width:1px"><img src="../UserMods/misc_cleardot.gif" height=1 width=1></td>
<%
		rsTemp.MoveNext
		if rsTemp.EOF then
			Response.Write "<td>&nbsp;</td><td>&nbsp;</td>"
			exit sub
		end if
%>
		<td nowrap valign=top align=center class="CPhomeImg">
			<%call prodImage()%>
		</td>
		<td valign=top class="CPhomeDesc" style="width:50%">
			<%call prodDetail()%>
		</td>
	</tr>
	<tr><td colspan=5><img src="../UserMods/misc_cleardot.gif" height=1 width=1></td></tr>
<%
end sub
'**********************************************************************
'Writes the product detail
'**********************************************************************
sub prodDetail()
%>
	<b class="CPprodDesc"><%=rsTemp("description")%></b><br><br>
	<span class="CPprodDescLong"><%=trim(rsTemp("descriptionLong"))%></span><br><br>
<%
	'Show pricing if required for this product
	if not(pHidePricingZero=-1 and rsTemp("Price")=0) then
	
		'Assign pricing info to local variables for easier use.
		dim listPrice, price
		listPrice = rsTemp("listPrice")
		price     = rsTemp("price")
	
		'Show Extended layout.
		if listViewLayout = 1 then
			if listPrice > Price then
				Response.Write "<span class=CPprodLPriceT>" & langGenListPrice & ":</span> <del class=CPprodLPriceV>" & pCurrencySign & moneyS((listPrice)) & "</del><br>"
			end if
			Response.Write "<b class=CPprodPriceT>" & langGenOurPrice & ":</b> <b class=CPprodPriceV>" & pCurrencySign & moneyS(Price) & "</b>"
			if (listPrice - Price) > 0 then
				Response.Write "<br><span class=CPprodSPriceT>" & langGenYouSave & ":</span> <span class=CPprodSPriceV>" & pCurrencySign & moneyS((listPrice-Price)) & " (" & formatNumber((((listPrice-Price)/listPrice)*100),0) & "%)</span>"
			end if
		'Show Classic layout.
		else
			Response.Write "<b class=CPprodPriceT>" & langGenOurPrice & ":</b> <b class=CPprodPriceV>" & pCurrencySign & moneyS(Price) & "</b>"
		end if
		Response.Write "<br><br>"
		
	end if
	
	'Show Extended layout.
	if listViewLayout = 1 then

		'Free Shipping?
		if UCase(rsTemp("noShipCharge")) = "Y" and len(trim(rsTemp("fileName")&"")) = 0 then
			Response.Write "<b class=CPfreeShipMsg>" & langGenFreeShipping & "</b><br>"
		end if
											
		'In stock, Out of stock
		if pShowStockView = -1 then
			if pHideAddStockLevel = -1 then
				Response.Write "<b class=CPinStock>" & langGenInStock & "</b><br>"
			else
				if rsTemp("stock") > pHideAddStockLevel then	
					Response.Write "<b class=CPinStock>" & langGenInStock & "</b><br>"
				else
					Response.Write "<b class=CPoutStock>" & langGenOutStock & "</b><br>"
				end if
			end if
		end if
		
	end if

end sub
'**********************************************************************
'Writes the code to display the product image and a link
'**********************************************************************
sub prodImage()
	if len(trim(rsTemp("smallImageUrl")&"")) <> 0 then
%>
		<a href="prodView.asp?idproduct=<%=rsTemp("idProduct")%>"><img src="<%=pImagesDir & rsTemp("smallImageUrl")%>" border=0 alt="<%=rsTemp("description")%>"></a><br>
<%
	else
%>
		<br><b class="CPnoImgT"><%=langGenNoImage%></b><br>
<%
	end if
%>
	<br><a href="prodView.asp?idproduct=<%=rsTemp("idProduct")%>"><%=langGenViewMore%></a><br>
<%
end sub
%>