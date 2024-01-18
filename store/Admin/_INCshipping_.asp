<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Shipping Rates Maintenance Functions Tabs
' Product  : CandyPress eCommerce Administration
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
sub shipTabs(activeTab)
%>
	<table border=0 cellpadding=0 cellspacing=0>
		<tr>
		
			<td align=center valign=middle bgcolor=<%=shipTabCol(activeTab,"OV")%> nowrap>
				&nbsp;&nbsp;<a href="SA_ship.asp">Overview</a>&nbsp;&nbsp;
			</td>
			
			<td width=1 bgcolor=#808080><img width=1 height=20 alt=""></td>

			<td align=center valign=middle bgcolor=<%=shipTabCol(activeTab,"CM")%> nowrap>
				&nbsp;&nbsp;<a href="SA_shipMet.asp">Store Methods</a>&nbsp;&nbsp;
			</td>
			
			<td width=1 bgcolor=#808080><img width=1 height=20 alt=""></td>
			
			<td align=center valign=middle bgcolor=<%=shipTabCol(activeTab,"CR")%> nowrap>
				&nbsp;&nbsp;<a href="SA_shipRate.asp?resetCookie=1">Store Rates</a>&nbsp;&nbsp;
			</td>
			
			<td width=1 bgcolor=#808080><img width=1 height=20 alt=""></td>
			
			<td align=center valign=middle bgcolor=<%=shipTabCol(activeTab,"UP")%> nowrap>
				&nbsp;&nbsp;<a href="SA_shipUPS.asp">UPS Online</a>&nbsp;&nbsp;
			</td>
			
			<td width=1 bgcolor=#808080><img width=1 height=20 alt=""></td>
			
			<td align=center valign=middle bgcolor=<%=shipTabCol(activeTab,"US")%> nowrap>
				&nbsp;&nbsp;<a href="SA_shipUSPS.asp">USPS Online</a>&nbsp;&nbsp;
			</td>
			
			<td width=1 bgcolor=#808080><img width=1 height=20 alt=""></td>
			
			<td align=center valign=middle bgcolor=<%=shipTabCol(activeTab,"CP")%> nowrap>
				&nbsp;&nbsp;<a href="SA_shipCP.asp">Canada Post</a>&nbsp;&nbsp;
			</td>
			
			<td width=1 bgcolor=#808080><img width=1 height=20 alt=""></td>
			
		</tr>
	</table>

	<table width=100% cellpadding=2 cellspacing=0 border=0>
		<tr>
			<td bgcolor=#DDDDCC align=left nowrap>&nbsp;</td>
		</tr>
	</table>

	<br>
<%
end sub

function shipTabCol(activeTab,currTab)
	if activeTab = currTab then
		shipTabCol = "#DDDDCC"
	else
		shipTabCol = "#EEEEEE"
	end if
end function
%>