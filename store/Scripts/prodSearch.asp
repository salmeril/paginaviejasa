<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Advanced product search form
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
dim idProduct
dim description
dim price

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

%>
<!--#include file="../UserMods/_INCtop_.asp"-->

<!-- Outer Table Cell -->
<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<!-- Heading -->
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr><td valign=middle class="CPpageHead">
		<b><%=langGenAdvancedSearch%></b>
	</td></tr>
</table>

<br>

<!-- Main Table -->
<table border="0" cellspacing="0" cellpadding="5" width="100%">
	<tr><td>
	 
	 	<form METHOD="POST" name="advSearch" action="prodList.asp">

			<%=langGenSearchKeywords%><br>
			<input type="text" name="strSearch" size="30"><br><br>

			<table border=0 cellspacing=0 cellpadding=0>
				<tr>
					<td nowrap><%=langGenSearchMinPrice%></td>
					<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td nowrap><%=langGenSearchMaxPrice%></td>
				</tr>
				<tr>
					<td nowrap><input type="text" name="strSearchMin" size="10" maxlength="10"></td>
					<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td nowrap><input type="text" name="strSearchMax" size="10" maxlength="10"></td>
				</tr>
			</table><br>

			<%=langGenSearchCat%><br>
			<select name="strSearchCat">
				<option value="0"><%=langGenAllCategories%></option>
<%
				mySQL = "SELECT   a.idCategory, a.categoryDesc " _
					  & "FROM     categories a, categories_products b " _
					  & "WHERE    a.idCategory = b.idCategory " _
					  & "GROUP BY a.idCategory, a.categoryDesc " _
					  & "ORDER BY a.categoryDesc "
				set rsTemp = openRSexecute(mySQL)
				do while not rsTemp.EOF
%>
					<option value="<%=rsTemp("idCategory")%>"><%=rsTemp("categoryDesc")%></option>
<%
					rsTemp.MoveNext
				loop
				call closeRS(rsTemp)
%>
			</select><br><br>

			<input type="SUBMIT" name="Submit" value="<%=langGenSearch%>">

		</form>

	</td></tr>
</table>

<!-- End Outer Table Cell -->
</td></tr></table>

<br>

<!--#include file="../UserMods/_INCbottom_.asp"-->

<%
call closedb()
%>
