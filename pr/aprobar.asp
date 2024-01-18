<%@LANGUAGE="VBSCRIPT"%><%
' *** Update Record: construct a sql update statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_tableName = "ordenes"
  MM_tableCol = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_fields = "fechaaprobado,fechaaprobado,#,none,NULL,aprobado,aprobado,',none,''"
  MM_redirectPage = "aprobar1.asp"

  ' create the sql update statement
  MM_updateStr = "update " & MM_tableName & " set "
  MM_fieldsArray = Split(MM_fields, ",")
  For i = LBound(MM_fieldsArray) To UBound(MM_fieldsArray) Step 5
    FormVal = CStr(Request.Form(MM_fieldsArray(i)))
    Delim = MM_fieldsArray(i+2)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_fieldsArray(i+3)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_fieldsArray(i+4)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fieldsArray)) Then
      MM_updateStr = MM_updateStr & ","
    End If
    MM_updateStr = MM_updateStr & MM_fieldsArray(i+1) & " = " & FormVal
  Next
  MM_updateStr = MM_updateStr & " where " & MM_tableCol & " = " & MM_recordId

  ' finish the sql and execute it
  Set MM_updateCmd = Server.CreateObject("ADODB.Command")
  MM_updateCmd.ActiveConnection = "dsn=5084.quotes;"
  MM_updateCmd.CommandText = MM_updateStr
  MM_updateCmd.Execute

  ' redirect with URL parameters
  If (MM_redirectPage = "") Then
    MM_redirectPage = CStr(Request("URL"))
  End If
  If (InStr(1, MM_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_redirectPage = MM_redirectPage & "?" & Request.QueryString
  End If
  Call Response.Redirect(MM_redirectPage)
End If
%> 
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
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF">
<table border="0" width="100%">
  <tr> 
    <td height="46">&nbsp;</td>
    <td height="46"> 
      <div align="left"></div>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> 
      <p>&nbsp; 
      <form name="form1" method="post" action="<%=MM_editAction%>">
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Orden #:</b> 
          <%=(rsordenes.Fields.Item("ID").Value)%></font> 
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Fecha:</b> 
          <%=(rsordenes.Fields.Item("fecha").Value)%></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
          </font> 
        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> <b>Nombre:</b> 
          <%=(rsordenes.Fields.Item("nombre").Value)%> 
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Apellido:</b> <%=(rsordenes.Fields.Item("apellido").Value)%><br>
          <b>Email:</b> <%=(rsordenes.Fields.Item("email").Value)%></font> 
        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> <b>Direcci&oacute;n:</b> 
          <%=(rsordenes.Fields.Item("direccion").Value)%> 
          <br>
          <b>Ciudad:</b> <%=(rsordenes.Fields.Item("ciudad").Value)%> 
          &nbsp;&nbsp;&nbsp;<b>Estado:</b> <%=(rsordenes.Fields.Item("estado").Value)%><br>
          <b>C&oacute;digo</b>: <%=(rsordenes.Fields.Item("codigo").Value)%> 
          &nbsp;&nbsp;&nbsp;<b>Tel&eacute;fono:</b> <%=(rsordenes.Fields.Item("telefono").Value)%> 
          </font> 
        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Forma 
          de pago:</b> <%=(rsordenes.Fields.Item("tarjeta").Value)%> 
          <br>
          </font> 
        <p><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Numero 
          de tarjeta:</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
          <%=(rsordenes.Fields.Item("numero").Value)%><br>
          <b> Fecha de vencimiento:</b> <%=(rsordenes.Fields.Item("expiracion").Value)%></font> 
        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Productos:</b><br>
          <%=(rsordenes.Fields.Item("productos").Value)%> 
          </font> 
        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Total:</b> 
          $<%=(rsordenes.Fields.Item("total").Value)%> 
          + $10 (gastos de envio)</font> 
        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Comentarios:</b><br>
          <%=(rsordenes.Fields.Item("comentarios").Value)%></font> 
          <input type="hidden" name="fechaaprobado" value="<%= date() %>">
          <input type="hidden" name="aprobado" value="si">
        <p> <%Session("svid")= (rsordenes.Fields.Item("id").Value) %> 
          <input type="submit" name="Submit" value="Aprobar">
          <input type="hidden" name="MM_recordId" value="<%= rsordenes.Fields.Item("ID").Value %>">
          <input type="hidden" name="MM_update" value="true">
          &nbsp;&nbsp;&nbsp;<a href="ordenview.asp">Cancelar</a> 
      </form>
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
        </font><font face="Verdana, Arial, Helvetica, sans-serif"> </font> 
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

