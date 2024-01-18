<% 
	' Generic interface to Jugos table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://208.159.3.50/asp/genericdb/main.htm"
	Session("dbTitle") = "Jugos"
	Session("dbType") = "Access"
	Session("dbConn") = "medicina"
	Session("dbRs") = "jugos"
	Session("dbKey") = 1
	Session("dbOrder") = 2
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "11000"
	Session("dbDispView") =     "11111"
	Session("dbDispEdit") =     "21111"
	Session("dbSearchFields") = "01000"
	Session("dbURLfor4") = 4
	Session("dbWhere") = ""
	Session("dbDebug") = 1
	Session("dbCanEdit") = 1
	Session("dbCanAdd") = 1
	Session("dbCanDelete") = 1
	Session("dbConfirmDelete") = 1
	Session("dbViewPage") = Request.ServerVariables("PATH_INFO")
	Response.Redirect Session("dbGenericPath") & "GenericList.asp"
%>
<html>

<head>
<title></title>
</head>

<body>
</body>
</html>
