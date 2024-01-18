<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level3.inc"-->

<html>
<body>

<center>
<table border="1" bgcolor="#c0c0c0">
<form action="update.asp?method=Add" method="Post">
<tr><td><b>Username</b></td><td><input type="text" name="username" size="10"></td></tr>
<tr><td><b>Password</b></td><td><input type="password" name="password" size="10"></td></tr>
<tr><td><b>Clearance Level</b> (0 - 3)</td>
<td>
<select name="level">
<option value="0">0
<option value="1">1
<option value="2">2	
<option value="3">Admin
</select>
</td></tr>
<tr><td><b>Expiration Date</b></td><td><input type="text" name="expdate" size="10" value="<%=DateAdd("yyyy", 1, Date)%>"></td></tr>
<tr><td bgcolor="#000000"><input type="submit" value="Add New Account"></td><td bgcolor="#c0c0c0">&nbsp;</tr>
</form>
</table>
</center>

<%
SQL = "Select ID, UserName, [PassWord], Clearance, ExpireDate From Login Order By ID"
Set RS = MyConn.Execute(SQL)

Response.Write "<center>"

While Not RS.EOF
  Response.Write "<form name=""Update"" method=""Post"">"
  Response.Write "<table border=""1"" bgcolor=""#c0c0c0"">"

  %>
  <tr><td><b>Username</b></td><td><b>Password</b></td><td><b>Level</b></td><td><b>Expiration Date</b></td></tr>
  <tr><td><input type="hidden" name="id" value="<%=RS("ID")%>"></td></tr>
  <tr>
  <td><input type="text" name="username" size="10" value="<%=RS("UserName")%>"></td>
  <td><input type="text" name="password" size="10" value="<%=RS("PassWord")%>"></td>
  <td><input type="text" name="level" size="1" value="<%=RS("Clearance")%>"></td>
  <td><input type="text" name="expdate" size="10" value="<%=RS("ExpireDate")%>"></td>
  <td bgcolor="#c0c0c0"><input type="submit" value="Update" onClick="this.form.action='update.asp?method=Edit';"></td>
  <td bgcolor="#c0c0c0"><input type="submit" value="Delete" onClick="this.form.action='update.asp?method=Delete';"></td>
  </tr>
  <%
  Response.Write "</table>"
  Response.Write "</form>"
  RS.MoveNext
Wend

Response.Write "</center>"    

CleanUp(RS)

Response.Write "<p><center><a href=""utility.asp?method=abandon""><b>Log Off</b></a></center>"
%>

</body>
</html>
