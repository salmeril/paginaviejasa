<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : UPS Online Shipping Rates
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
dim UPSactive
dim UPSAccessID
dim UPSUserID
dim UPSPassword
dim UPSfromZip
dim UPSfromCntry
dim UPSpickupType
dim UPSpackType
dim UPSshipCode
dim UPSweightUnit
dim UPSallRates

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
	case lCase("UPSactive")
		UPSactive			= rs("configVal")
	case lCase("UPSAccessID")
		UPSAccessID			= rs("configVal")
	case lCase("UPSUserID")
		UPSUserID			= rs("configVal")
	case lCase("UPSPassword")
		UPSPassword			= rs("configVal")
	case lCase("UPSfromZip")
		UPSfromZip			= rs("configVal")
	case lCase("UPSfromCntry")
		UPSfromCntry		= rs("configVal")
	case lCase("UPSpickupType")
		UPSpickupType		= rs("configVal")
	case lCase("UPSpackType")
		UPSpackType			= rs("configVal")
	case lCase("UPSshipCode")
		UPSshipCode			= rs("configVal")
	case lCase("UPSweightUnit")
		UPSweightUnit		= rs("configVal")
	end select

	rs.MoveNext
loop
call closeRS(rs)

'Close Database
call closedb()
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Shipping - UPS Online Shipping Rates</font></b>
	<br><br>
</P>
<%
'Page Tabs
call shipTabs("UP")

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
		the shipping rates returned by UPS are not adjusted for 
		these items. Always enter a few typical test orders to 
		verify that you are getting the shipping rate results 
		you want.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Note 2.</td>
	<td valign=top>
		The rates returned by this routine will always 
		be in the currency of the shipment's originating country. In 
		other words, if you are shipping from Canada to the US, the 
		rates returned will be in Canadian dollars. Similarly, if you 
		are shipping from the US to Canada, the rates returned will be 
		in US dollars.
	</td>
</tr>
</table>

<br><span class="textBlockHead">Step-By-Step :</font><br>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr>
	<td nowrap valign=top>Step 1.</td>
	<td valign=top>
		XML - Your web server must be able to communicate with the 
		UPS servers via Microsoft's XML components. Checking for XML 
		components --&gt; 
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
		UPS ACCOUNT - Before you can use the UPS Online Shipping 
		Rates routine, you will need to 
		<a href="http://www.ec.ups.com/ecommerce/gettools/gtools_intro.html" target="_blank">register</a> 
		with UPS as an "End User" in 
		order to use their "Online Tools" products. This is a fairly 
		elaborate process, but as long as you follow their instructions, 
		you will eventually obtain an 'Access Key'. The 'Access Key' is 
		required to use the routine.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 4.</td>
	<td valign=top>
		COUNTRY CODES - You will need to ensure that the country 
		codes assigned to each country in your store are valid ISO 
		country codes (by default they are). UPS only accepts the 
		ISO format.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 5.</td>
	<td valign=top>
		SHIPMENT WEIGHT - The weight of your products must be 
		entered into your database as Pounds or Kilograms (non-US only). 
		You can have fractions (eg. 0.5 Pounds, 1.89 Pounds).
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 6.</td>
	<td valign=top>
		CONFIGURE - Complete the form below to configure your store 
		for UPS Online Rates.<br><br>
		<table border=0 cellspacing=0 cellpadding=5 class="blockInBlock">
		<form method="post" action="SA_shipUPS_exec.asp" name="form1">
			<tr>
				<td nowrap valign=top><b>Active?</b></td>
				<td valign=top>
					<input type=checkbox name=UPSactive value="Y" <%if UPSactive="Y" then Response.Write "checked" end if%>>
					<span class="fieldHelp">
						(Check box to activate Online UPS rates)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>UPS Access Key</b></td>
				<td valign=top>
					<input type=text name=UPSAccessID id=UPSAccessID size=20 value="<%=UPSAccessID%>"> 
					<span class="fieldHelp">
						(Value is case sensitive)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>UPS User ID</b></td>
				<td valign=top>
					<input type=text name=UPSUserID id=UPSUserID size=20 value="<%=UPSUserID%>"> 
					<span class="fieldHelp">
						(Value is case sensitive)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>UPS Password</b></td>
				<td valign=top>
					<input type=text name=UPSPassword id=UPSPassword size=20 value="<%=UPSPassword%>"> 
					<span class="fieldHelp">
						(Value is case sensitive)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Country Code</b></td>
				<td valign=top>
					<input type=text name=UPSfromCntry id=UPSfromCntry size=2 maxlength=2 value="<%=UPSfromCntry%>">
					<span class="fieldHelp">
						(Origination ISO Country Code)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Zip/Post Code</b></td>
				<td valign=top>
					<input type=text name=UPSfromZip id=UPSfromZip size=10 value="<%=UPSfromZip%>">
					<span class="fieldHelp">
						(Origination Zip/Post Code)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Select UPS Service</b></td>
				<td valign=top>
					<select name=UPSshipCode id=UPSshipCode size=1>
						<option value=""   <%=checkMatch(UPSshipCode,"")  %>>Show rates for all available services</option>
						<option value="01" <%=checkMatch(UPSshipCode,"01")%>>Next Day Air</option>
						<option value="02" <%=checkMatch(UPSshipCode,"02")%>>2nd Day Air</option>
						<option value="03" <%=checkMatch(UPSshipCode,"03")%>>Ground</option>
						<option value="07" <%=checkMatch(UPSshipCode,"07")%>>Worldwide Express</option>
						<option value="08" <%=checkMatch(UPSshipCode,"08")%>>Worldwide Expedited</option>
						<option value="11" <%=checkMatch(UPSshipCode,"11")%>>Standard</option>
						<option value="12" <%=checkMatch(UPSshipCode,"12")%>>3-Day Select</option>
						<option value="13" <%=checkMatch(UPSshipCode,"13")%>>Next Day Air Saver</option>
						<option value="14" <%=checkMatch(UPSshipCode,"14")%>>Next Day Air Early AM</option>
						<option value="54" <%=checkMatch(UPSshipCode,"54")%>>Worldwide Express Plus</option>
						<option value="59" <%=checkMatch(UPSshipCode,"59")%>>2nd Day Air AM</option>
						<option value="65" <%=checkMatch(UPSshipCode,"65")%>>Express Saver</option>
					</select><br>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Pickup Type</b></td>
				<td valign=top>
					<select name=UPSpickupType id=UPSpickupType size=1>
						<option value="01" <%=checkMatch(UPSpickupType,"01")%>>Daily Pickup</option>
						<option value="03" <%=checkMatch(UPSpickupType,"03")%>>Customer Counter</option>
						<option value="06" <%=checkMatch(UPSpickupType,"06")%>>One Time Pickup</option>
						<option value="07" <%=checkMatch(UPSpickupType,"07")%>>On Call Air</option>
						<option value="19" <%=checkMatch(UPSpickupType,"19")%>>Letter Center</option>
						<option value="20" <%=checkMatch(UPSpickupType,"20")%>>Air Service Center</option>				
					</select><br>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Package Type</b></td>
				<td valign=top>
					<select name=UPSpackType id=UPSpackType size=1>
						<option value="02" <%=checkMatch(UPSpackType,"02")%>>Package</option>
						<option value="00" <%=checkMatch(UPSpackType,"00")%>>Unknown</option>
						<option value="01" <%=checkMatch(UPSpackType,"01")%>>UPS Letter</option>
						<option value="03" <%=checkMatch(UPSpackType,"03")%>>UPS Tube</option>
						<option value="04" <%=checkMatch(UPSpackType,"04")%>>UPS Pak</option>
						<option value="21" <%=checkMatch(UPSpackType,"21")%>>UPS Express Box</option>
						<option value="24" <%=checkMatch(UPSpackType,"24")%>>UPS 25KG Box</option>
						<option value="25" <%=checkMatch(UPSpackType,"25")%>>UPS 10KG Box</option>
					</select><br>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Weight Type</b></td>
				<td valign=top>
					<select name=UPSweightUnit id=UPSweightUnit size=1>
						<option value="LBS" <%=checkMatch(UPSweightUnit,"LBS")%>>Pounds</option>
						<option value="KGS" <%=checkMatch(UPSweightUnit,"KGS")%>>Kilograms</option>
					</select> 
					<span class="fieldHelp">
						(US customers can only specify Pounds)<br>
					</span>
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
