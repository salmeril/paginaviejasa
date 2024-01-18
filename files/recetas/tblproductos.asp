<% 
	' Generic interface to Productos table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://208.159.3.50/asp/genericdb/main.htm"
	Session("dbTitle") = "Productos Naturales"
	Session("dbType") = "Access"
	Session("dbConn") = "medicina"
	Session("dbRs") = "productos"
	Session("dbKey") = 1
	Session("dbOrder") = 2
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "11000000"
	Session("dbDispView") =     "11111111"
	Session("dbDispEdit") =     "21111111"
	Session("dbSearchFields") = "01000000"
	Session("dbURLfor7") = 7
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
