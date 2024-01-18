<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Canada Post Online Shipping Rates
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
const adminLevel = 0
%>
<!--#include file="../Scripts/_INCconfig_.asp"-->
<!--#include file="_INCsecurity_.asp"-->
<!--#include file="_INCappDBConn_.asp"-->
<!--#include file="_INCshipping_.asp"-->
<!--#include file="../UserMods/_INClanguage_.asp"-->
<!--#include file="../Scripts/_INCappFunctions_.asp"-->
<%
'Form variables
dim CPactive
dim CPmerchantID
dim CPfromZip
dim CPsizeL
dim CPsizeW
dim CPsizeH

'Database variables
dim mySQL, cn, rs

'Work Fields
dim xmlTest
'*************************************************************************

'Open Database
call openDb()

'Get current configuration settings from database
mySQL = "SELECT configVar, configVal " _
	  & "FROM   storeAdmin " _
	  & "WHERE  adminType = 'S'"
set rs = openRSexecute(mySQL)
do while not rs.EOF

	select case trim(lCase(rs("configVar")))
	case lCase("CPactive")
		CPactive			= rs("configVal")
	case lCase("CPmerchantID")
		CPmerchantID		= rs("configVal")
	case lCase("CPfromZip")
		CPfromZip			= rs("configVal")
	case lCase("CPsizeL")
		CPsizeL				= rs("configVal")
	case lCase("CPsizeW")
		CPsizeW				= rs("configVal")
	case lCase("CPsizeH")
		CPsizeH				= rs("configVal")
	end select

	rs.MoveNext
loop
call closeRS(rs)

'Close Database
call closedb()
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Shipping - Canada Post Online Shipping Rates</font></b>
	<br><br>
</P>
<%
'Page Tabs
call shipTabs("CP")

if len(trim(Request.QueryString("msg"))) > 0 then
%>
	<font color=red><%=Request.QueryString("msg")%></font>
	<br><br>
<%
end if
%>

<span class="textBlockHead">Notes :</font><br>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr>
	<td nowrap valign=top>Note 1.</td>
	<td valign=top>
		Stores that sell irregularly shaped or oversized 
		items requiring additional shipping charges may find that 
		the shipping rates returned by Canada Post are not adjusted 
		for these items. Always enter a few typical test orders to 
		verify that you are getting the shipping rate results 
		you want.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Note 2.</td>
	<td valign=top>
		The rates returned by this routine will always be in Canadian 
		dollars.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Note 3.</td>
	<td valign=top>
		Canada Post currently only provide shipping for packages that 
		originate from within Canada.
	</td>
</tr>
</table>

<br><span class="textBlockHead">Step-By-Step :</font><br>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr>
	<td nowrap valign=top>Step 1.</td>
	<td valign=top>
		XML - Your web server must be able to communicate with the 
		Canada Post servers via Microsoft's XML components. Checking 
		for XML components --&gt; 
<%
		on error resume next
		set xmlTest = Server.CreateObject("Microsoft.XMLDOM")
		if err.number = 0 then
			set xmlTest = server.Createobject("MSXML2.ServerXMLHTTP")
			if err.number = 0 then
				Response.Write "<font color=green>COMPONENTS INSTALLED</font>"
			else
				err.Clear
				set xmlTest = server.Createobject("MSXML2.ServerXMLHTTP.4.0")
				if err.number = 0 then
					Response.Write "<font color=green>COMPONENTS INSTALLED</font>"
				else
					Response.Write "<font color=red>MSXML2.SERVERXMLHTTP NOT FOUND</font>"
				end if
			end if
		else
			Response.Write "<font color=red>MICROSOFT.XMLDOM NOT FOUND</font>"
		end if
		set xmlTest = nothing
		on error goto 0
%>
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 2.</td>
	<td valign=top>
		IIS 5.0 - You will need to have IIS 5.0 (or later) to use the 
		shipping routine. You currently have 
		<b><%=Request.ServerVariables("SERVER_SOFTWARE")%></b> 
		installed. 
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 3.</td>
	<td valign=top>
		CANADA POST PROFILE - Before you can use the Canada Post 
		Online Shipping Rates routine, you will need to create a 
		"Shipping Profile" by sending a request to 
		<a href="mailto:sellonline@canadapost.ca">sellonline@canadapost.ca</a>. 
		They will then guide you through the rest of the process. You 
		will initially receive a "test" account when the profile is 
		created, but you can request them to move the account to 
		production immediately because the module is already tested.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 4.</td>
	<td valign=top>
		COUNTRY CODES - You will need to ensure that the country 
		codes assigned to each country in your store are valid ISO 
		country codes (by default they are). Also, shipping to the 
		US requires that US state codes are the standard 2 letter 
		abbreviation (by default they are).
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 5.</td>
	<td valign=top>
		SHIPMENT WEIGHT - The weight of your products must be 
		entered into your database as Kilograms. You can have 
		fractions (eg. 1.89 Kilograms).
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 6.</td>
	<td valign=top>
		CONFIGURE - Complete the form below to configure your store 
		for Canada Post Online Rates.<br><br>
		<table border=0 cellspacing=0 cellpadding=5 class="blockInBlock">
		<form method="post" action="SA_shipCP_exec.asp" name="form1">
			<tr>
				<td nowrap valign=top><b>Active?</b></td>
				<td valign=top>
					<input type=checkbox name=CPactive value="Y" <%if CPactive="Y" then Response.Write "checked" end if%>>
					<span class="fieldHelp">
						(Check box to activate Canada Post rates)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>User ID</b></td>
				<td valign=top>
					<input type=text name=CPmerchantID id=CPmerchantID size=25 value="<%=CPmerchantID%>"> 
					<span class="fieldHelp">
						(Value is case sensitive)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Post Code</b></td>
				<td valign=top>
					<input type=text name=CPfromZip id=CPfromZip size=6 maxlength=6 value="<%=CPfromZip%>">
					<span class="fieldHelp">
						(Origination Post Code)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=middle><b>Avg. Parcel Size<br>(Centimeters)</b></td>
				<td valign=middle>
					Length <input type=text name=CPsizeL id=CPsizeL size=2 value="<%=CPsizeL%>"> 
					Width  <input type=text name=CPsizeW id=CPsizeW size=2 value="<%=CPsizeW%>"> 
					Height <input type=text name=CPsizeH id=CPsizeH size=2 value="<%=CPsizeH%>">
				</td>
			</tr>
			<tr>
				<td colspan=2 align=center>
					<br><input type=submit name=submit1  id=submit1  value="Submit"><br><br>
				</td>
			</tr>
		</form>
		</table>
		<br><br>
	</td>
</tr>
</table>

<!--#include file="_INCfooter_.asp"--> 
