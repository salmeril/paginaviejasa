<% 
	' Generic interface to Productos table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://www.nuestramedicina.com/asp/genericdb/main.htm"
	Session("dbTitle") = "Productos Naturales"
	Session("dbType") = "Access"
	Session("dbConn") = "7598.medicina"
	Session("dbRs") = "productos"
	Session("dbKey") = 1
	Session("dbOrder") = 2
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "110000000"
	Session("dbDispView") =     "111111110"
	Session("dbDispEdit") =     "211111110"
	Session("dbSearchFields") = "010000000"
	Session("dbURLfor8") = 8
	Session("dbWhere") = ""
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
