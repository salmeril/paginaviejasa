
<% 
	' Generic interface to Inventarios table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://www.nuestramedicina.com/asp/genericdb/admin.htm"
	Session("dbTitle") = "Inventarios"
	Session("dbType") = "Access"
	Session("dbConn") = "5084.inventarios"
	Session("dbRs") = "all"
	Session("dbKey") = 1
	Session("dbOrder") = 2
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "010000000000000000"
	Session("dbDispView") =     "111111111111111111"
	Session("dbDispEdit") =     "111111111111111111"
	Session("dbSearchFields") = "010000000000000000"
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
