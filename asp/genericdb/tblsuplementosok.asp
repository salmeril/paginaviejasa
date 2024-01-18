<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level3.inc"-->

<% 
	' Generic interface to Suplementos table. 
	Session("dbGenericPath") = "/asp/genericdb/"
	Session("dbExitPage") = "http://www.nuestramedicina.com/asp/genericdb/admin.htm"
	Session("dbTitle") = "Suplementos"
	Session("dbType") = "Access"
	Session("dbConn") = "5084.suplementos"
	Session("dbRs") = "suplementos ok"
	Session("dbKey") = 1
	Session("dbOrder") = 3
	Session("dbRecsPerPage") = 20
	Session("dbDispList") =     "001000000000000000000"
	Session("dbDispView") =     "111111111111111111111"
	Session("dbDispEdit") =     "211111111111111111111"
	Session("dbSearchFields") = "001000000000000000000"
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
