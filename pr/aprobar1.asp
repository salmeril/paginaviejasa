<%@LANGUAGE="VBSCRIPT"%> <%

Dim rsordenes__strID
rsordenes__strID = "0"
if(Session("svID") <> "") then rsordenes__strID = Session("svID")

%> <%
set rsordenes = Server.CreateObject("ADODB.Recordset")
rsordenes.ActiveConnection = "dsn=5084.quotes;"
rsordenes.Source = "SELECT *  FROM ordenes  WHERE ID = " + Replace(rsordenes__strID, "'", "''") + ""
rsordenes.CursorType = 0
rsordenes.CursorLocation = 2
rsordenes.LockType = 3
rsordenes.Open
rsordenes_numRows = 0
%> 
<html>
<head>
<title>Confirmacion de aprobacion</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF">
<table width="75%" border="0" cellspacing="10" cellpadding="10">
  <tr align="center"> 
    <td> 
      <p><%
id =(rsordenes.Fields.Item("ID").Value)
email = (rsordenes.Fields.Item("email").Value)

Set objMail = CreateObject("CDONTS.Newmail")
objMail.From ="drarmoza@hotmail.com"
objMail.To = email
objMail.Subject = "Orden"
objMail.Body = "Su orden ha sido enviado. Gracias."
objMail.Send
  %></p>
      <p><font face="Arial, Helvetica, sans-serif"><b><font face="Verdana, Arial, Helvetica, sans-serif">Email 
        de aprobacion ha sido enviado</font></b></font></p>
      <p><font face="Verdana, Arial, Helvetica, sans-serif"><a href="ordenall.asp">volver</a></font></p>
    </td>
  </tr>
</table>
</body>
</html>

