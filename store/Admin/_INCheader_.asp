<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Header Page
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
<HTML>

<HEAD>
	<TITLE>CandyPress eCommerce Administration</TITLE>
	<STYLE type="text/css">
	<!--
	BODY, B, TD, P, UL, LI, INPUT, SELECT, TEXTAREA
	{
		COLOR: #333333; FONT-FAMILY: Verdana, Arial, helvetica; FONT-SIZE: 8pt;
	}
	.errMsg			 {COLOR: red;}
	.linkclass       {COLOR: #333333; TEXT-DECORATION: none;}
	.linkclass:hover {COLOR: #972A24; TEXT-DECORATION: underline;}
	.spacer          {COLOR: #333333; FONT-FAMILY: Arial, Verdana, Helvetica; FONT-SIZE: 8px;}
	.fieldHelp       {COLOR: #800000; FONT-FAMILY: Verdana, Arial, Helvetica; FONT-SIZE: 8pt;}
	.textBlockHead   {COLOR: #800000; FONT-FAMILY: Verdana, Arial, Helvetica; FONT-SIZE: 9pt; FONT-WEIGHT: BOLD;}
	.navigationBar   {BACKGROUND-COLOR: #EEEEEE; BORDER: 1px solid #888888;}
	.textBlock       {BACKGROUND-COLOR: #EEEEEE; BORDER: 1px solid #CCCCCC;}
	.blockInBlock    {BACKGROUND-COLOR: #F5F5F5; BORDER: 1px dashed #888888;}
	.findTable       {BACKGROUND-COLOR: #EEEEEE; BORDER: 1px solid #CCCCCC;}
	.listTable       {BACKGROUND-COLOR: #EEEEEE;}
	.listRowTop      {BACKGROUND-COLOR: #DDDDCC; BORDER-TOP: 1px solid #888888; BORDER-BOTTOM: 1px solid #888888;}
	.listRowHead     {BACKGROUND-COLOR: #DDDDDD;}
	.listRowBot      {BACKGROUND-COLOR: #DDDDCC; BORDER-TOP: 1px solid #888888; BORDER-BOTTOM: 1px solid #888888;}
	-->
	</STYLE>
</HEAD>

<BODY>

<TABLE border=0 cellpadding=0 cellspacing=0 width=100%>
	<TR>
		<TD nowrap align=left valign=middle bgcolor="#EEEEEE">
			<b><font size=2>CandyPress eCommerce Administration</font></b>
		</TD>
		<TD nowrap align=right valign=middle bgcolor="#EEEEEE">
<%
			'Show logon level			
			if trim(session(storeID & "adminLoggedOn")) = "" then
				Response.Write "<font color=red>Logged Off</font>"
			else
				Response.Write "<font color=green>Logged On (" & session(storeID & "adminLoggedOn") & ")</font>"
			end if
%>
			&nbsp;&nbsp;
			[ 
			<a href="default.asp" class="linkclass">Home</a> | 
			<a href="logon_exec.asp?action=logoff" class="linkclass">Logoff</a> 
			] 
		</TD>
	</TR>
	<TR>
		<TD colspan=2 align=center valign=middle>
			<img src="x_pixel.gif" width="100%" height="1" align="absMiddle" vspace=4>
		</TD>
	</TR>
</TABLE>

<TABLE border=0 cellpadding=5 cellspacing=0 width=100%>
	<TR>
		<TD align=left valign=top width="100" nowrap class="navigationBar">
			<img src="x_cleardot.gif" border=0 width=100 height=1><br>
			<a href="utilities.asp" class="linkclass">Setup & Utilities</a><br><span class="spacer"><br></span>
<%
			if UCase(StoreAdminInstalled) = "Y" then
%>
			<a href="SA_cat.asp" class="linkclass">Categories</a><br><span class="spacer"><br></span>
			<a href="SA_opt.asp" class="linkclass">Options</a><br><span class="spacer"><br></span>
			<a href="SA_prod.asp?resetCookie=1" class="linkclass">Products</a><br><span class="spacer"><br></span>
			<a href="SA_rev.asp?resetCookie=1" class="linkclass">Reviews</a><br><span class="spacer"><br></span>
			<a href="SA_loc.asp?resetCookie=1" class="linkclass">Locations & Tax</a><br><span class="spacer"><br></span>
			<a href="SA_ship.asp" class="linkclass">Shipping</a><br><span class="spacer"><br></span>
			<a href="SA_order.asp?resetCookie=1" class="linkclass">Orders</a><br><span class="spacer"><br></span>
			<a href="SA_cust.asp?resetCookie=1" class="linkclass">Customers</a><br><span class="spacer"><br></span>
			<a href="SA_disc.asp?resetCookie=1" class="linkclass">Discounts</a><br><span class="spacer"><br></span>
			<a href="SA_stats.asp" class="linkclass">Statistics</a><br><span class="spacer"><br></span>
			<a href="SA_news.asp" class="linkclass">Newsletters</a><br><span class="spacer"><br></span>
			<br><br><br><br><br><br><br>
<%
			else
%>
			<br><br><br><br><br><br><br><br><br>
			<br><br><br><br><br><br><br><br><br>
			<br><br><br><br><br><br><br><br><br>
<%
			end if
%>
		</TD>
		<TD align=left valign=top width="10">
			&nbsp;
		</TD>
		<TD align=left valign=top width="100%">
