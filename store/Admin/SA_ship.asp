<%@ Language=VBScript %>
<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Shipping - Main Page
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
<!--#include file="_INCshipping_.asp"-->
<%
'Work Fields
dim mySQL

'*************************************************************************

%>

<!--#include file="_INCheader_.asp"-->

<P align=left>
	<b><font size=3>Shipping Maintenance</font></b>
	<br><br>
</P>
<%
'Page Tabs
call shipTabs("OV")
%>
<span class="textBlockHead">Overview :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr>
		<td>
			This is an overview of the shipping options available 
			to you for calculating shipping rates for your orders. 
			Shipping is fairly simple to set up, as long as you 
			understand what tools are available, and how to use these 
			tools to their fullest extent. We therefore suggest that 
			you take the time to read the notes on this page to get 
			a better idea of how shipping rates are implemented. First 
			of all, you need to familiarize yourself with the concept 
			of <b>Store</b>, <b>Online</b> and <b>Custom</b> shipping 
			rate calculations.
		</td>
	</tr>
</table>

<br>

<span class="textBlockHead">Store Rates :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr>
		<td>
			<b>Overview</b> - Store Rates refer to the 
			standard shipping rate calculation mechanism that comes 
			with the software. The rates themselves are defined by you, 
			and entered into the database using the supplied  
			functions. Store rates are, by their very nature, not meant 
			to be precise representations of actual rates. For example, 
			if you create Store Rates for 'UPS Ground', the rates you 
			enter will invariably differ from the actual rates charged 
			by UPS. This is because UPS factors in many variables that 
			are not necessarily available or known to you at the time 
			of calculation.<br><br>
			
			<b>Zones</b> - The first determination you need to make is 
			how you want to group countries and/or states together 
			as shipping zones. Each zone is treated as a 
			single shipping destination for the purpose of calculating a 
			shipping rate. For example, you can group all the countries 
			in Europe together 
			as a single zone, in which case the same shipping rate will 
			be calculated for all the countries in Europe. Or you may 
			want to group all the western US states into a single zone. 
			Grouping geographical areas greatly simplifies the entry 
			of shipping rates into your database. Be aware though that, 
			the more countries or states you group into a single zone, the 
			greater the difference between your rates and the actual 
			shipping rates charged by the shipping company will be.<br><br>
			
			<b>Shipping Method</b> - This is used to group together a set 
			of shipping rates (eg. 'UPS Ground', 'FedEx 2 Day Air', etc.). 
			Each shipping method can apply to one or more shipping zone(s), 
			which in turn can each have their own unique set of shipping 
			rates. Therefore, you can have a unique set of shipping rates 
			for each method/zone combination. As an example, this allows 
			you to only display 'UPS Ground' shipping rates for US based 
			orders, and 'FedEx International' for non-US destinations.<br><br>
			
			<b>Shipping Rates</b> - Once you have determined which shipping 
			methods you are going to use, and which zones they apply to, 
			you can start entering the detail rates for 
			each method/zone combination. Each shipping rate that you enter 
			must have a weight or price range. This refers to the 
			total weight and price of the order. So if you want to add 
			$10.00 to orders that weigh between 1 and 5 pounds, you will create 
			a shipping rate similar to this :<br><br>
			
			<i>Rate Type = 'Weight'; From = '1.00'; To = '5.00'; Add Amount = '10'</i><br><br>
			
			You can also choose to enter a percentage, instead 
			of a fixed amount. Naturally, the greater the weight or price 
			range, the less accurate the calculated rate will be. For example 
			UPS charges a certain amount for each pound of weight. To closely 
			resemble this, you will have to enter a shipping rate record for 
			each pound : <br><br>
			
			<i>Rate Type = 'Weight'; From = '0.00'; To = '1.00'; Add Amount = '1'</i><br>
			<i>Rate Type = 'Weight'; From = '1.01'; To = '2.00'; Add Amount = '2'</i><br>
			...<br><br>
			
			This will of course require a large number of shipping rate 
			records to be entered. Instead, it may be better to average 
			the rates out over multiple pounds :<br><br>
			
			<i>Rate Type = 'Weight'; From = '0.00'; To = '5.00'; Add Amount = '3'</i><br>
			<i>Rate Type = 'Weight'; From = '5.01'; To = '9.00'; Add Amount = '7'</i><br>
			...<br><br>

			This may mean that you will be short a few cents on some orders, and 
			over on others, but you would have greatly eased the process 
			of entering your shipping rates.<br><br>
			
			<b>Unit of Weight</b> - The measurement that you use to enter 
			weight can be anything - Pounds, Kilograms, Ounces, Grams, etc. 
			What is important is that you USE THE SAME MEASUREMENT FOR ALL 
			WEIGHT RELATED ENTRIES THROUGHOUT THE ENTIRE STORE. So if you 
			have entered the weight of your products as Pounds, you must 
			enter the weight ranges for your shipping rates in Pounds.<br><br>
			
		</td>
	</tr>
</table>

<br>

<span class="textBlockHead">Online Rates :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr>
		<td>
			<b>Overview</b> - Online Rates are realtime shipping rates 
			obtained directly from a shipping company by accessing 
			their servers over the internet. The rates returned by 
			these routines can be used to supplement your Store Rates, 
			or they can be used exclusively for 
			calculating shipping rates for your store. Each Online 
			Rate routine has it's own set of requirements and limitations, 
			so be sure to check the additional help for the online 
			routine you want to use.<br><br>
		</td>
	</tr>
</table>

<br>

<span class="textBlockHead">Custom Rates :</span><br>
<table border=0 cellspacing=0 cellpadding=5 width="100%" class="textBlock">
	<tr>
		<td>
			<b>Overview</b> - If you want to develop your own shipping 
			rate routines, or enhance the supplied Store Rates routine, 
			then Custom Rates are the way to do it. To develop your 
			own Custom Rate routines you will need to have some ASP 
			programming knowledge.<br><br>
			
			<b>User Exit File</b> - To overcome the problem of 
			upgrading the software without overwriting your Custom 
			Rate routines, we have provided a special file called a 
			User Exit into which all Custom Rate routines must be 
			placed. This file is located at <b>UserMods/_INCship_.asp</b> 
			and can be edited with any good text based editor, such 
			as Notepad. Any future upgrades to the main software 
			folders will therefore leave your changes intact.<br><br>
			
		</td>
	</tr>
</table>

<!--#include file="_INCfooter_.asp"--> 
