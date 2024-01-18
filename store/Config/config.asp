<%
'*********************************************************************
' Function : Database and Security Settings
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*********************************************************************
dim connString

'---------------------------------------------------------------------
' Database Connection String.
'
' EXAMPLE : DSN-less connection for SQL Server
' connString="driver={SQL Server};server=SERVER_NAME;uid=USER_ID;pwd=PASSWORD;database=DB_NAME"
'
' EXAMPLE : DSN-less connection for MS Access
' connString="driver={Microsoft Access Driver (*.mdb)};DBQ=C:\FULL_PATH_TO_DB\DB_NAME.mdb;"
'
' EXAMPLE : DSN connection (all databases)
' connString="dsn=NAME_OF_YOUR_DSN"
'---------------------------------------------------------------------
connString = "dsn=5084.store"


'---------------------------------------------------------------------
' Database Type (0=Access ; 1=SQL Server)
'---------------------------------------------------------------------
const dbType = 0


'---------------------------------------------------------------------
' Administrator UserID and Password
'---------------------------------------------------------------------
const adminUser = "admin"
const adminPass = "admin"


'---------------------------------------------------------------------
' RC4Key used for En/De-cription. CHANGE ONLY ONCE!
'
' This is the Encryption Key used to Encrypt/Decrypt Passwords and
' Credit Card Numbers on your database. It is recommended that you
' write your key down and keep it in a safe place in case you need it.
'
' The key must be :
' 1. Exactly 30 characters in length
' 2. Consist only of Alpha-Numeric characters (a-z,A-Z,0-9)
' 3. No special characters
' 4. The key is case sensitive
' 5. Example : "nw8vwstsldi6tsadhsd389705vj5hd"
'
' WARNING : Once this key has been entered, you can NOT change it
' again. If you do, you will no longer be able to decrypt existing
' information that was encrypted with the old key.
'---------------------------------------------------------------------
const rc4Key = "armonianaturalarmonianaturalar"


'---------------------------------------------------------------------
' Store ID. If you are going to be hosting multiple stores on the same 
' web (or virtual web), you should assign a different store ID to each 
' of them to prevent sessions from being shared. The storeID can be 
' any combination of alpha (a-z) and numeric (0-9) characters with a
' maximum of 10 characters (eg. "toystore"), and a minimum of 1
' character. No spaces are allowed.
'---------------------------------------------------------------------
const storeID = "store"


'---------------------------------------------------------------------
' Put Store in Demo Mode? (Y/N)
'---------------------------------------------------------------------
const demoMode = "N"


'---------------------------------------------------------------------
' Is CandyPress Store Admin installed. (Y/N)
'---------------------------------------------------------------------
const StoreAdminInstalled = "Y"


'-- ADDED - Version 1.8 ----------------------------------------------
' Put Store in Debug Mode? (Y/N)
'---------------------------------------------------------------------
const debugMode = "Y"


'-- ADDED - Version 2.2 ----------------------------------------------
' Lock the Database. (Y/N)
'---------------------------------------------------------------------
const dbLocked = "N"


'-- ADDED - Version 2.2 ----------------------------------------------
' Additional login for general non-administrator staff. This login 
' will NOT allow acces to the setup and configuration utilities.
'---------------------------------------------------------------------
const nonAdminUser = ""
const nonAdminPass = ""

%>
