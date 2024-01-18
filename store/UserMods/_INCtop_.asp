<%
'*************************************************************************
' Function : HTML/ASP "header" code which is added to the top of 
'          : every page. This page serves as an example and can be 
'          : customized to suit your store.
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
%>
<!--#include file="_INCtools_.asp"-->
<html>
<head>

	<title>Online Store</title>
	
	<meta name="keywords"    content="online,store,shop,cart,storefront">
	<meta name="description" content="Online Store">
	
	<link rel="stylesheet" type="text/css" href="../UserMods/_INCstyles_.css">
	
</head>

<body>

<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr>
		
		<!-- Company Logo -->
		<td align="left" valign="middle">
			<img alt="<%=pCompany%>" border=0 src="../UserMods/logo.gif">
		</td>
    
		<!-- Search Form -->
		<form action="<%=urlNonSSL%>prodList.asp" method="post" id="searchBox" name="searchBox">
		<td align="left" valign="middle">
			<input name="strSearch" size="20" value="<%=langGenSearch%>" align=absmiddle onFocus="javascript:if (this.value=='<%=langGenSearch%>') {this.value='';}"> 
			<input src="../UserMods/butt_go.gif" type="image" border="0" name="SubSearch" align="absmiddle"><br>
			<a href="<%=urlNonSSL%>prodSearch.asp"><%=langGenAdvancedSearch%></a><br>
		</td>
		</form>
		
		

		<!-- Cart Summary -->
		<td align="right" valign="middle">
			<table border=0 cellSpacing=0 cellPadding=1 width=140 class="CPbox2">
				<tr>
					<td nowrap align=center colspan=2 class="CPbox2H"><b><%=langGenShoppingCart%></b></td>
				</tr>
				<tr>
					<td nowrap class="CPbox2B"><%=langGenQty%></td>
			        <td noWrap class="CPbox2B"><b><%=cartQty(idOrder)%></b></td>
			    </tr>
				<tr>
					<td noWrap class="CPbox2B"><%=langGenTotal & "(" & pcurrencySign & ")"%></td>
					<td noWrap class="CPbox2B"><b><%=moneyS(cartTotal(idOrder,0))%></b></td>
				</tr>
			</table>
		</td>

	</tr>
</table>

<!-- Top Navigation Bar -->
<img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=4><br>
<table border="0" cellspacing="0" cellpadding="2" width="100%" class="CPbox1">
	<tr>
		<td align=left valign=middle width="50%"> 
			&nbsp;<a href="<%=urlNonSSL%>default.asp"><%=langGenHome%></a> 
			<img src="../UserMods/misc_dot.gif" border=0> 
			<a href="<%=urlNonSSL%>prodList.asp"><%=langGenAllCategories%></a> 
			<img src="../UserMods/misc_dot.gif" border=0> 
			<a href="<%=urlNonSSL%>contactUs.asp"><%=langGenContactUsHdr%></a> 
		</td>
		<td align=right valign=middle width="50%">
			<a href="<%=urlNonSSL%>05_Gateway.asp?action=logon"><%=langGenAccount%></a> 
			<img src="../UserMods/misc_dot.gif" border=0> 
			<a href="<%=urlNonSSL%>cart.asp"><%=langGenCart%></a> 
			<img src="../UserMods/misc_dot.gif" border=0> 
			<a href="<%=urlNonSSL%>05_Gateway.asp?action=checkout"><%=langGenCheckout%></a>&nbsp;
		</td>
	</tr>
</table>
<img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=4><br>

<!-- Main Shopping Cart Area -->
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
	<tr>
	
		<td align="left" valign="top" width="135">
		
			<!-- Categories -->
			<img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=5><br>
			<table border=0 cellSpacing=0 cellPadding=5 width=135 class="CPbox2">
				<tr>
					<td nowrap align=center class="CPbox2H">
						<b><%=langGenCategories%></b>
					</td>
				</tr>
				<tr>
					<td class="CPbox2B">
						<a href="<%=urlNonSSL%>prodList.asp"><%=langGenAllCategories%></a>&nbsp;<br>
						<a href="<%=urlNonSSL%>prodList.asp?special=Y"><%=langGenSpecials%></a>&nbsp;<br>
						<%=showFeaturedCat()%>
					</td>
			    </tr>
			</table>
			
			<!-- What's New -->
			<img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=8><br>
			<table border=0 cellSpacing=0 cellPadding=5 width=135 class="CPbox2">
				<tr>
					<td nowrap align=center class="CPbox2H">
						<b><%=langGenNewProd%></b>
					</td>
				</tr>
				<tr>
					<td class="CPbox2B">
						<%=showNewProd(5)%>
					</td>
			    </tr>
			</table>

			<!-- Top Sellers -->
			<img src="../UserMods/misc_cleardot.gif" border=0 width=1 height=8><br>
			<table border=0 cellSpacing=0 cellPadding=5 width=135 class="CPbox2">
				<tr>
					<td nowrap align=center class="CPbox2H">
						<b><%=langGenTopSellers%></b>
					</td>
				</tr>
				<tr>
					<td class="CPbox2B">
						<%=showTopSell(5)%>
					</td>
			    </tr>
			</table>

		</td>
		
		<!-- Spacer -->
		<td style="width:10px">
			<img src="../UserMods/misc_cleardot.gif" border=0 width=10 height=1>
		</td>
		
		<!-- Shopping Cart -->
		<td align="center" valign="top" width="100%">
			<img src="../UserMods/misc_cleardot.gif" height=6 width=1><br>
