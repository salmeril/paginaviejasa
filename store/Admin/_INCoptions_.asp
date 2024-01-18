<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Options Maintenance Functions Tabs
' Product  : CandyPress eCommerce Administration
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************
sub optTabs(activeTab)
%>
	<table border=0 cellpadding=0 cellspacing=0>
		<tr>
		
			<td align=center valign=middle bgcolor=<%=optTabCol(activeTab,"OP")%> nowrap width=120>
				<a href="SA_opt.asp">Options</a>
			</td>
			
			<td width=1 bgcolor=#808080><img width=1 height=20 alt=""></td>

			<td align=center valign=middle bgcolor=<%=optTabCol(activeTab,"OG")%> nowrap width=120>
				<a href="SA_optGrp.asp">Option Groups</a>
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

function optTabCol(activeTab,currTab)
	if activeTab = currTab then
		optTabCol = "#DDDDCC"
	else
		optTabCol = "#EEEEEE"
	end if
end function
%>