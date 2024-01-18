<%@LANGUAGE="VBSCRIPT"%> 
<%

Dim rsordenes__MMColParam
rsordenes__MMColParam = "0"
if(Request.QueryString("ID") <> "") then rsordenes__MMColParam = Request.QueryString("ID")

%> <%
set rsordenes = Server.CreateObject("ADODB.Recordset")
rsordenes.ActiveConnection = "dsn=5084.quotes;"
rsordenes.Source = "SELECT * FROM ordenes WHERE ID = " + Replace(rsordenes__MMColParam, "'", "''") + ""
rsordenes.CursorType = 0
rsordenes.CursorLocation = 2
rsordenes.LockType = 3
rsordenes.Open
rsordenes_numRows = 0
%> 
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF">
<table border="0" width="100%">
  <tr> 
    <td height="46">&nbsp;</td>
    <td height="46"> 
      <div align="left"><a href="ordenall.asp">ver todos</a> </div>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> 
      <p> <font face="Verdana, Arial, Helvetica, sans-serif">Orden #: <b><%=(rsordenes.Fields.Item("ID").Value)%></b></font> 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Fecha: <b><%=(rsordenes.Fields.Item("fecha").Value)%></b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        </font> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> Nombre: 
        <b><%=(rsordenes.Fields.Item("nombre").Value)%></b> 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Apellido: <b><%=(rsordenes.Fields.Item("apellido").Value)%></b><br>
        Email: <b><%=(rsordenes.Fields.Item("email").Value)%></b></font> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> Direcci&oacute;n: 
        <b><%=(rsordenes.Fields.Item("direccion").Value)%></b> 
        <br>
        Ciudad: <b><%=(rsordenes.Fields.Item("ciudad").Value)%></b><br>
        Estado: <b><%=(rsordenes.Fields.Item("estado").Value)%></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        C&oacute;digo: <b><%=(rsordenes.Fields.Item("codigo").Value)%></b><br>
        </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Tel&eacute;fono: 
        <b><%=(rsordenes.Fields.Item("telefono").Value)%></b> 
        </font> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Forma de 
        pago: <b><%=(rsordenes.Fields.Item("tarjeta").Value)%></b> 
        <br>
        </font> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Numero de 
        tarjeta:</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <b><%=(rsordenes.Fields.Item("numero").Value)%></b><br>
        Fecha de vencimiento: <b><%=(rsordenes.Fields.Item("expiracion").Value)%></b></font>
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Productos:<br>
        <b><%=(rsordenes.Fields.Item("productos").Value)%></b> 
        </font> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Total: <b>$<%=(rsordenes.Fields.Item("total").Value)%></b> 
        + $10 (gastos de envio)</font> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Comentarios:<br>
        <b><%=(rsordenes.Fields.Item("comentarios").Value)%></b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
        </font> 
    </td>
  </tr>
  <tr> 
    <td align="right" height="16"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif"> </font> 
    </td>
    <td height="16"> 
      <p>&nbsp; 
    </td>
  </tr>
  <tr> 
    <td align="right">&nbsp;</td>
    <td height="5">&nbsp;</td>
  </tr>
</table>
</body>
</html>
