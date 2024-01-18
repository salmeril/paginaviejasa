<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Newsletters and Mailing List
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
const adminLevel = 1
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<%
'Newsletters
dim idNews
dim newsBookmark
dim newsSubj
dim newsBody

'Database
dim mySQL, cn, rs, rs2

'*************************************************************************

'Open Database Connection
call openDB()

'Store Configuration
if loadConfig() = false then
	response.redirect "sysMsg.asp?errMsg=" & server.URLEncode("Could not load Store Configuration settings.")
end if

'Check for most recent newsletter
mySQL="SELECT TOP 1 " _
	& "       idNews,newsBookmark,newsSubj,newsBody " _
    & "FROM   newsletters " _
    & "ORDER BY idNews "
set rs = openRSexecute(mySQL)
if not rs.eof then
	newsBody	= trim(rs("newsBody"))
	idNews		= rs("idNews")
	newsBookmark= trim(rs("newsBookmark"))
	newsSubj	= trim(rs("newsSubj"))
end if
call closeRS(rs)

'Close database connection
call closedb()

%>
<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Newsletters and Mailing Lists</font></b>
	<br><br>
</P>

<form method="post" action="SA_news_exec.asp" name="newsletters">

	<table border=0 cellspacing=0 cellpadding=3 width="450" class="textBlock">
		<tr>
			<td colspan=2><b><font color=#800000>1. Select Customers</font></b></td>
		</tr>
		<tr>
			<td><input type="radio" name="custType" value="A"></td>
			<td nowrap><b>All Customers</b></td>
		</tr>
		<tr>
			<td><input type="radio" name="custType" value="I"></td>
			<td nowrap><b>Opt-In Customers Only</b></td>
		</tr>
		<tr>
			<td><input type="radio" name="custType" value="O"></td>
			<td nowrap><b>Opt-Out Customers Only</b></td>
		</tr>
		<tr>
			<td><input type="checkbox" name="custPaid" value="Y"></td>
			<td nowrap><b>Only include customers with PAID orders?</b></td>
		</tr>
		<tr><td colspan=2>&nbsp;</td></tr>
		<tr>
			<td colspan=2><b><font color=#800000>2. Select Action</font></b></td>
		</tr>
		<tr>
			<td><input type="radio" name="action" value="D"></td>
			<td nowrap><b>Display Customers</b></td>
		</tr>
		<tr>
			<td><input type="radio" name="action" value="F"></td>
			<td nowrap><b>Download Customers</b></td>
		</tr>
		<tr>
			<td><input type="radio" name="action" value="E"></td>
			<td nowrap><b>Email Newsletters to Customers</b> (Complete Form Below)</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>
			
				<table border=0 cellspacing=0 cellpadding=5 width="95%" class="blockInBlock">
					<tr>
						<td nowrap>From</td>
						<td><b><%=pEmailSales%></b></td>
					</tr>
					<tr>
						<td nowrap>Subject</td>
						<td><input type="text" name="newsSubj" size="40" maxlength="255" value="<%=server.HTMLEncode(newsSubj & "")%>"></td>
					</tr>
					<tr>
						<td nowrap>Message</td>
						<td><textarea name="newsBody" cols="50" rows="9"><%=server.HTMLEncode(newsBody & "")%></textarea>&nbsp;&nbsp;&nbsp;</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>
							<input type="checkbox" name="contType" value="1"> 
							Send as HTML Email<br>
							<font size=1 color="#800000">
								Check this box if you want the 'Message' above to 
								be interpreted as HTML by your customer's email 
								reader. 
							</font>
						</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>
							<input type="checkbox" name="newsPreview" value="Y"> 
							Preview Mode?<br>
							<font size=1 color="#800000">
								Preview mode will ignore the customer list, and 
								send one copy of the newsletter to <%=pEmailSales%>. 
							</font>
						</td>
					</tr>
<%
					if len(newsBookmark) > 0 then
%>
					<tr>
						<td>&nbsp;</td>
						<td>
							<input type="checkbox" name="newsBookmark" value="<%=newsBookmark%>"> 
							Start from <b><%=newsBookmark%></b> ?<br>
							<font size=1 color="#800000">
								The previous newsletter ended before all the 
								emails were sent out. Check the box to start the 
								current newsletter at the last known successful 
								email of the previous newsletter.
							</font>
						</td>
					</tr>
<%
					end if
%>
					<tr><td colspan=2>&nbsp;</td></tr>
					<tr>
						<td colspan=2>
							<b>Note: </b> Each newsletter is 
							individually emailed to each of the 
							customers on the list. In cases where 
							you have several thousand customers on 
							your mailing list and a slow email 
							server, you could wait for an hour or 
							even more before the next page is 
							displayed.
						</td>
					</tr>
					<tr><td colspan=2>&nbsp;</td></tr>
				</table>
			</td>
		</tr>
		<tr><td colspan=2>&nbsp;</td></tr>
		<tr>
			<td colspan=2 align=center>
				<input type="hidden" name="idNews"  value="<%=idNews%>">
				<input type="submit" name="submit1" value="Submit">
			</td>
		</tr>
		<tr><td colspan=2>&nbsp;</td></tr>
	</table>
	
</form>

<!--#include file="_INCfooter_.asp"-->
