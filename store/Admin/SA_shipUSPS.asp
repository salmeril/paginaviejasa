<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : USPS Online Shipping Rates
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
dim USPSactive
dim USPSUserID
dim USPSPassword
dim USPSfromZip
dim USPSservice
dim USPSintNtl
dim USPSsize
dim USPSmachinable

'Database variables
dim mySQL, cn, rs

'Work Fields
dim serviceArray
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
	case lCase("USPSactive")
		USPSactive			= rs("configVal")
	case lCase("USPSUserID")
		USPSUserID			= rs("configVal")
	case lCase("USPSPassword")
		USPSPassword		= rs("configVal")
	case lCase("USPSfromZip")
		USPSfromZip			= rs("configVal")
	case lCase("USPSservice")
		USPSservice			= rs("configVal")
	case lCase("USPSintNtl")
		USPSintNtl			= rs("configVal")
	case lCase("USPSsize")
		USPSsize			= rs("configVal")
	case lCase("USPSmachinable")
		USPSmachinable		= rs("configVal")
	end select

	rs.MoveNext
loop
call closeRS(rs)

'Close Database
call closedb()

'Create service array
serviceArray = split(USPSservice,",")
%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Shipping - USPS Online Shipping Rates</font></b>
	<br><br>
</P>
<%
'Page Tabs
call shipTabs("US")

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
		the shipping rates returned by USPS are not adjusted for 
		these items. Always enter a few typical test orders to 
		verify that you are getting the shipping rate results 
		you want.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Note 2.</td>
	<td valign=top>
		The rates returned by this routine will always be in US dollars.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Note 3.</td>
	<td valign=top>
		USPS currently only provide shipping for packages that originate 
		from within the United States.
	</td>
</tr>
</table>

<br><span class="textBlockHead">Step-By-Step :</font><br>

<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
<tr>
	<td nowrap valign=top>Step 1.</td>
	<td valign=top>
		XML - Your web server must be able to communicate with the 
		USPS servers via Microsoft's XML components. Checking for XML 
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
		USPS ACCOUNT - Before you can use the USPS Online Shipping 
		Rates routine, you will need to 
		<a href="http://www.uspsprioritymail.com/et_regcert.html" target="_blank">register</a> 
		with USPS to use their "Web Tools" product. You will initially 
		receive a "test" account, but you can go ahead and request a 
		"live" account as the routines have already been tested.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 4.</td>
	<td valign=top>
		SHIPMENT WEIGHT - The weight of your products must be 
		entered into your database as Pounds. The weight will be 
		rounded to the nearest Pound when calculating rates.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 5.</td>
	<td valign=top>
		US COUNTRY CODE - Check that the Country Code for the United 
		States is "US" using the Locations function. This is the 
		default value.
	</td>
</tr>
<tr>
	<td nowrap valign=top>Step 6.</td>
	<td valign=top>
		CONFIGURE - Complete the form below to configure your store 
		for USPS Online Rates.<br><br>
		<table border=0 cellspacing=0 cellpadding=5 class="blockInBlock">
		<form method="post" action="SA_shipUSPS_exec.asp" name="form1">
			<tr>
				<td nowrap valign=top><b>Active?</b></td>
				<td valign=top>
					<input type=checkbox name=USPSactive value="Y" <%if USPSactive="Y" then Response.Write "checked" end if%>>
					<span class="fieldHelp">
						(Check box to activate Online USPS rates)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>USPS User Name</b></td>
				<td valign=top>
					<input type=text name=USPSUserID id=USPSUserID size=20 value="<%=USPSUserID%>"> 
					<span class="fieldHelp">
						(Value is case sensitive)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>USPS Password</b></td>
				<td valign=top>
					<input type=text name=USPSPassword id=USPSPassword size=20 value="<%=USPSPassword%>"> 
					<span class="fieldHelp">
						(Value is case sensitive)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Zip/Post Code</b></td>
				<td valign=top>
					<input type=text name=USPSfromZip id=USPSfromZip size=10 value="<%=USPSfromZip%>">
					<span class="fieldHelp">
						(Origination Zip/Post Code)
					</span><br>
				</td>
			</tr>
			<tr>
				<td nowrap valign=middle><b>Select US Services</b></td>
				<td valign=middle nowrap>
					<input type=checkbox name=USPSservice1 value="Express"     <%=checkArr(serviceArray,"Express","checked")    %>>Express &nbsp;
					<input type=checkbox name=USPSservice2 value="First Class" <%=checkArr(serviceArray,"First Class","checked")%>>First Class &nbsp;
					<input type=checkbox name=USPSservice3 value="Priority"    <%=checkArr(serviceArray,"Priority","checked")   %>>Priority &nbsp;
					<input type=checkbox name=USPSservice4 value="Parcel"      <%=checkArr(serviceArray,"Parcel","checked")     %>>Parcel &nbsp;
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Ship International?</b></td>
				<td valign=top>
					<select name=USPSintNtl id=USPSintNtl size=1>
						<option value="Y" <%=checkMatch(USPSintNtl,"Y")%>>Yes</option>
						<option value="N" <%=checkMatch(USPSintNtl,"N")%>>No</option>
					</select>
					<span class="fieldHelp">
						(Allow shipping outside US?)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Package Size</b></td>
				<td valign=top>
					<select name=USPSsize id=USPSsize size=1>
						<option value="REGULAR"  <%=checkMatch(USPSsize,"REGULAR") %>>Regular</option>
						<option value="LARGE"    <%=checkMatch(USPSsize,"LARGE")   %>>Large</option>
						<option value="OVERSIZE" <%=checkMatch(USPSsize,"OVERSIZE")%>>Oversize</option>
					</select>
					<span class="fieldHelp">
						(See USPS Documentation)<br>
					</span>
				</td>
			</tr>
			<tr>
				<td nowrap valign=top><b>Machinable?</b></td>
				<td valign=top>
					<select name=USPSmachinable id=USPSmachinable size=1>
						<option value="TRUE"  <%=checkMatch(USPSmachinable,"TRUE") %>>True</option>
						<option value="FALSE" <%=checkMatch(USPSmachinable,"FALSE")%>>False</option>
					</select>
					<span class="fieldHelp">
						(See USPS Documentation)<br>
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
<%
'*********************************************************************
'Check an array for a value and return requested string
'*********************************************************************
function checkArr(arr, strVal, strRet)
	dim i
	for i = 0 to Ubound(arr)
		if LCase(arr(i)) = LCase(strVal) then
			checkArr = strRet
			exit for
		end if
	next
end function
%>
