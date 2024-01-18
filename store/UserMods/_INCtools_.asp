<%
'*************************************************************************
' Function : HTML/ASP code functions and subroutines which can be 
'          : used to customize the look of the store's header.
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
'Display Featured Categories
'*************************************************************************
function showFeaturedCat()

	'Declare some variables
	dim mySQL, rsTemp, tempStr
	
	'Read Database
	mySQL = "SELECT idCategory,categoryDesc " _
		  & "FROM   categories " _
		  & "WHERE  categoryFeatured='Y' " _
		  & "ORDER BY categoryDesc "
	set rsTemp = openRSexecute(mySQL)
	do while not rsTemp.EOF
	
		'Build display string
		tempStr = tempStr _
			& "<a href=""" & urlNonSSL _
			& "prodList.asp?idCategory=" & rsTemp("idCategory") & """>" _
			& rstemp("categoryDesc") _
			& "</a><br>"
			
		'Next Record
		rsTemp.MoveNext
	
	loop
	call closeRS(rsTemp)

	'Return	
	showFeaturedCat = tempStr
 
end function

'*************************************************************************
'Display New Products
'*************************************************************************
function showNewProd(dispNum)

	'Declare some variables
	dim mySQL, rsTemp, tempStr, count
	
	'Read Database
	mySQL="SELECT TOP " & dispNum & " idProduct, description " _
	    & "FROM   products " _
	    & "WHERE  active = -1 " _
	    & "ORDER BY idProduct DESC "
	set rsTemp = openRSexecute(mySQL)
	do while not rsTemp.EOF
	
		'Increment counter
		count = count + 1
		
		'Build display string
		tempStr = tempStr _
			& count & ". <a href=""" & urlNonSSL _
			& "prodView.asp?idProduct=" & rsTemp("idProduct") & """>" _
			& rstemp("description") _
			& "</a><br>"
			
		'Next Record
		rsTemp.MoveNext
	
	loop
	call closeRS(rsTemp)

	'Return
	if isEmpty(tempStr) or isNull(tempStr) or len(trim(tempStr)) = 0 then
		showNewProd = langGenNotApplicable
	else
		showNewProd = tempStr
	end if
 
end function

'*************************************************************************
'Display Top Sellers
'*************************************************************************
function showTopSell(dispNum)

	'Declare some variables
	dim mySQL, rsTemp, tempStr, count
	
	'Read Database
	mySQL = "SELECT TOP " & dispNum & " a.idProduct, a.description " _
	      & "FROM   cartRows a, cartHead b " _
	      & "WHERE  a.idOrder = b.idOrder " _
	      & "AND   (b.orderStatus='1' OR b.orderStatus='2' OR b.orderStatus='7') " _
	      & "GROUP BY a.idProduct, a.description " _
	      & "ORDER BY SUM(a.quantity) DESC "
	set rsTemp = openRSexecute(mySQL)
	do while not rsTemp.EOF
	
		'Increment counter
		count = count + 1
		
		'Build display string
		tempStr = tempStr _
			& count & ". <a href=""" & urlNonSSL _
			& "prodView.asp?idProduct=" & rsTemp("idProduct") & """>" _
			& rstemp("description") _
			& "</a><br>"
			
		'Next Record
		rsTemp.MoveNext
	
	loop
	call closeRS(rsTemp)

	'Return
	if isEmpty(tempStr) or isNull(tempStr) or len(trim(tempStr)) = 0 then
		showTopSell = langGenNotApplicable
	else
		showTopSell = tempStr
	end if
 
end function
%>