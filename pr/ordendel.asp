<%@LANGUAGE="VBSCRIPT"%><%
' *** Delete Record: construct a sql delete statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_tableName = "ordenes"
  MM_tableCol = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_redirectPage = "ordenalladm.asp"

  ' create the delete sql statement
  MM_deleteStr = "delete from " & MM_tableName & " where " & MM_tableCol & " = " & MM_recordId

  ' finish the sql and execute it
  Set MM_deleteCmd = Server.CreateObject("ADODB.Command")
  MM_deleteCmd.ActiveConnection = "dsn=5084.quotes;"
  MM_deleteCmd.CommandText = MM_deleteStr
  MM_deleteCmd.Execute

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
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsrecomendados_numRows = rsrecomendados_numRows + Repeat1__numRows
%> <%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF">
<p>&nbsp; </p>
<table border="0" width="100%">
  <tr> 
    <td width="20">&nbsp;</td>
    <td height="5">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20">&nbsp;</td>
    <td height="30"> 
      <p><font face="Arial, Helvetica, sans-serif"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#330066">BORRAR</font></b></font> 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif"> </font> 
    </td>
    <td> 
      <form name="delete" method="post" action="<%=MM_editAction%>">
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Orden #:</b> 
          <%=(rsordenes.Fields.Item("ID").Value)%></font> 
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Fecha:</b> 
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
          <%=(rsordenes.Fields.Item("comentarios").Value)%><br>
          </font><font face="Verdana, Arial, Helvetica, sans-serif"></font> 
        <p>&nbsp;</p>
        <p> <font face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="submit" name="borrar" value="BORRAR">
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ordenalladm.asp">cancelar</a> 
          </font> </p>
        <input type="hidden" name="MM_recordId" value="<%= rsordenes.Fields.Item("ID").Value %>">
        <input type="hidden" name="MM_delete" value="true">
      </form>
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> </font> 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right">&nbsp;</td>
    <td height="5">&nbsp;</td>
  </tr>
</table>
<p>&nbsp; </p>
</body>
</html>

