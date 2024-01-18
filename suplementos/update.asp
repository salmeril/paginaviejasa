<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level3.inc"-->

<%
Dim Method

Method = Request.QueryString("method")

Select Case Method
  Case "Add"
    Add(MyConn)
  Case "Edit"
    Edit(MyConn)
  Case "Delete"
    Delete(MyConn)
End Select

'/////////////////////////////////////////////////////////////////////////////////

Sub Add(MyConn)

  Dim UserName, PassWord, Level, ExpDate, SQL

  UserName = Replace(Trim(Request.Form("username")), "'", "''")
  PassWord = Replace(Trim(Request.Form("password")), "'", "''")
  Level = Trim(Request.Form("level"))
  ExpDate = Trim(Request.Form("expdate"))

  If UserName = "" Or PassWord = "" Or Level = "" Or ExpDate = "" Then Response.Redirect "admin.asp"

  SQL = "Insert Into Login (UserName, [PassWord], Clearance, ExpireDate) Values('"&UserName&"', '"&PassWord&"', '"&Level&"', '"&ExpDate&"')"

  MyConn.Execute(SQL)

  CleanUp2()
  
  Response.Redirect "admin.asp"

End Sub

'////////////////////////////////////////////////////////////////////////////////////

Sub Edit(MyConn)
  
  Dim id, username, password, level, expdate
  
  id = CInt(Request.Form("id"))
  username = Replace(Request.Form("username"), "'", "''")
  password = Replace(Request.Form("password"), "'", "''")
  level = CInt(Request.Form("level"))
  expdate = Request.Form("expdate")

  SQL = "Update Login Set UserName = '"&username&"', [PassWord] = '"&password&"'"
  SQL = SQL & ", Clearance = "&level&", ExpireDate = '"&expdate&"' Where ID = "&id&""

  Set RS = MyConn.Execute(SQL)

  CleanUp2()

  Response.Redirect "admin.asp"

End Sub

'////////////////////////////////////////////////////////////////////////////////////

Sub Delete(MyConn)

  Dim id, SQL

  id = CInt(Request.Form("id"))

  SQL = "Delete * From Login Where ID = "&id&""
  MyConn.Execute(SQL)

  CleanUp2()
  
  Response.Redirect "admin.asp"

End Sub
%>
