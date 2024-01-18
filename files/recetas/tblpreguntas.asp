<% 
	' Generic interface to Preguntas table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://208.159.3.50/asp/genericdb/main.htm"
	Session("dbTitle") = "Preguntas Frequentes"
	Session("dbType") = "Access"
	Session("dbConn") = "medicina"
	Session("dbRs") = "preguntas"
	Session("dbKey") = 1
	Session("dbOrder") = 2
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "110000"
	Session("dbDispView") =     "111111"
	Session("dbDispEdit") =     "211111"
	Session("dbSearchFields") = "010000"
	Session("dbURLfor3") = 3
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
