<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : RC4 Encryption / Decryption routines
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************

'******************************************************************
'Encrypt/Decrypt text
'Note : Script is based on RC4 routines by Mike Shaffer with mods
'     : by Zivan van Zyl
'******************************************************************
Function EnDeCrypt(plaintxt, psw)

	'If plaintext is empty, return Empty String
	if isEmpty(plaintxt) or isNull(plaintxt) or plaintxt = "" then
		EnDeCrypt = ""
		exit function
	end if
	
	'If psw (RC4Key) is incorrect, stop everything and display an error.
	if len(psw) <> 30 then
		'Just in case
		Response.Clear
		'Show error
		Response.Redirect "sysMsg.asp?errMsg=" & server.URLEncode(langErrInvRC4Key)
		'Just in case
		exit function
	end if

	'Declare Variables
	dim sbox(255), key(255)
	dim temp, tempSwap, intLength
	dim a, b, i, j, k
	dim cipherby, cipher

	'Initialize some variables
	b = 0
	i = 0
	j = 0
   
	'Initialize sbox and key array
	intLength = len(psw)
	For a = 0 To 255
		key(a)   = asc(mid(psw, (a mod intLength)+1, 1))
		sbox(a)  = a
	next
	For a = 0 To 255
		b = (b + sbox(a) + key(a)) Mod 256
		tempSwap = sbox(a)
		sbox(a)  = sbox(b)
		sbox(b)  = tempSwap
	Next

	'Encrypt/Decrypt text
	For a = 1 To Len(plaintxt)
		i		 = (i + 1) Mod 256
		j		 = (j + sbox(i)) Mod 256
		temp	 = sbox(i)
		sbox(i)  = sbox(j)
		sbox(j)  = temp
		k		 = sbox((sbox(i) + sbox(j)) Mod 256)
	    cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
	    cipher	 = cipher & Chr(cipherby)
	Next

	EnDeCrypt = cipher

End Function
'******************************************************************
'Convert a String to Hex values
'******************************************************************
function Ascii2Hex(strTemp)
	if strTemp = "" or isNull(strTemp) then
		Ascii2Hex = ""
	else
		dim I
		for I = 1 to len(strTemp)
			Ascii2Hex = Ascii2Hex & right("00" & hex(asc(mid(strTemp,I,1))),2)
		next
	end if
end function
'******************************************************************
'Convert a Hex values to String
'******************************************************************
function Hex2Ascii(strTemp)
	if strTemp = "" or isNull(strTemp) then
		Hex2Ascii = ""
	else
		dim I
		for I = 1 to len(strTemp) step 2
			Hex2Ascii = Hex2Ascii & Chr(Eval("&H" & Mid(strTemp,I,2)))
		next
	end if
end function
%>