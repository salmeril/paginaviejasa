<% 
	' Generic interface to Noticias table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://www.nuestramedicina.com/novedades.htm"
	Session("dbTitle") = "Novedades"
	Session("dbType") = "Access"
	Session("dbConn") = "7598.medicina"
	Session("dbRs") = "noticias"
	Session("dbKey") = 1
	Session("dbDispList") =     "010"
	Session("dbDispView") =     "011"
	Session("dbDispEdit") =     "011"
	Session("dbWhere") = ""
	Session("dbCanEdit") = 1
	Session("dbCanAdd") = 1
	Session("dbCanDelete") = 1
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
