<%
'**********************************************************************
' Function : Userexit that can be used to create a customized 
'          : shipping routine. Should only be modified by someone
'          : with experience in ASP development. 
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'**********************************************************************
'1. The sole purpose of the code in this script must be to populate an 
'   existing two-dimensional array with shipping rates and descriptions. 
'   The array is already declared so there is no need to do so again.
'
'2. The array that must be populated is shipArray(100,2). The first 
'   dimension (0) must contain the rate and must have valid numeric 
'   data. The second dimension (1) contains the rate description and 
'   has a maximum length of 100 chars.
'
'3. Other fields that are available to you :
'   totalShipPrice     = Total price  for non-free shipping items
'   totalShipWeight    = Total weight for non-free shipping items
'   shippingLocCountry = Country code of shipping address
'   shippingLocState   = State code of shipping address
'   shippingZip        = Zip / Postal Code of shipping address
'   locShipZone        = Zone for shipping address
'
'**********************************************************************
%>

