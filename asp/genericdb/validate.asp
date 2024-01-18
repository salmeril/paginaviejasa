<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->

<html>
<body>

<%
UserName = Replace(Trim(Request.Form("username")), "'", "''")
PassWord = Replace(Trim(Request.Form("password")), "'", "''")

If UserName = "" OR PassWord = "" Then Response.Redirect "default.asp"

SQL = "Select ID, UserName, [PassWord], Clearance, ExpireDate From Login"
Set RS = MyConn.Execute(SQL)

While Not RS.EOF  
  If UserName = RS("UserName") And PassWord = RS("Password") Then
    If RS("ExpireDate") > Now() Then
      Session("allow") = True
      Session("clearance") = RS("Clearance")
      Level = RS("Clearance")
    Else
      Response.Redirect "utility.asp?method=expired"
    End If
  End If
  RS.MoveNext
Wend

CleanUp(RS)

If Session("allow") = True Then
  If Level = 3 Then Response.Redirect "admin.asp"
  If Level < 3 Then Response.Redirect "welcome.asp"
Else
  Response.Redirect "default.asp"
End If
%>

</body>
</html>
