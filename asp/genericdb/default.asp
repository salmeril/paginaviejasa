<%@Language=VBSCript%>
<%Response.Buffer=True%>

<html>
<body>

<form action="validate.asp" method="Post">
<input type="text" name="username" size="20"> UserName<br>
<input type="password" name="password" size="20"> Password<br>
<input type="submit" value="Log In">
</form>

<%
If Session("allow") = False Then
  Response.Write "You are not currently logged in."
Else
  Response.Write "You are currently logged in."
  Response.Write "<p><a href=""utility.asp?method=abandon"">Log Off</a>"
End If
%>

</body>
</html>
