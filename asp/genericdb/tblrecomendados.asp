<% 
	' Generic interface to Radios table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://www.nuestramedicina.com/asp/genericdb/main.htm"
	Session("dbTitle") = "recomendados"
	Session("dbType") = "Access"
	Session("dbConn") = "5084.quotes"
	Session("dbRs") = "recomendados"
	Session("dbKey") = 1
	Session("dbOrder") = 2
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "11000"
	Session("dbDispView") =     "11111"
	Session("dbDispEdit") =     "21111"
	Session("dbSearchFields") = "01000"
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
