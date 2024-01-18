<% 
	' Generic interface to Radios table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://www.nuestramedicina.com/asp/genericdb/main.htm"
	Session("dbTitle") = "Radios"
	Session("dbType") = "Access"
	Session("dbConn") = "7598.medicina"
	Session("dbRs") = "radios"
	Session("dbKey") = 1
	Session("dbOrder") = 2
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "1100"
	Session("dbDispView") =     "1110"
	Session("dbDispEdit") =     "2110"
	Session("dbSearchFields") = "0100"
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
