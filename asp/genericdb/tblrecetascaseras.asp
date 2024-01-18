<% 
	' Generic interface to Preguntas table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://www.nuestramedicina.com/asp/genericdb/main.htm"
	Session("dbTitle") = "Recetas Caseras"
	Session("dbType") = "Access"
	Session("dbConn") = "7598.medicina"
	Session("dbRs") = "recetascaseras"
	Session("dbKey") = 1
	Session("dbOrder") = 2
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "110000"
	Session("dbDispView") =     "111110"
	Session("dbDispEdit") =     "211110"
	Session("dbSearchFields") = "010000"
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
