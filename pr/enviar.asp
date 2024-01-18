<%@LANGUAGE="VBSCRIPT"%>

<%Session("svid")= request.form("id") %> 
<%
' *** Update Record: construct a sql update statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_tableName = "recomendados"
  MM_tableCol = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_fields = "id,ID,none,none,NULL,enviados,enviado,',none,''"
  MM_redirectPage = "view.asp"

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

Dim rsrecomendados__MMColParam
rsrecomendados__MMColParam = "0"
if(Request.QueryString("ID") <> "") then rsrecomendados__MMColParam = Request.QueryString("ID")

%> <%
set rsrecomendados = Server.CreateObject("ADODB.Recordset")
rsrecomendados.ActiveConnection = "dsn=5084.quotes;"
rsrecomendados.Source = "SELECT * FROM recomendados WHERE ID = " + Replace(rsrecomendados__MMColParam, "'", "''") + ""
rsrecomendados.CursorType = 0
rsrecomendados.CursorLocation = 2
rsrecomendados.LockType = 3
rsrecomendados.Open
rsrecomendados_numRows = 0
%> 
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsrecomendados_numRows = rsrecomendados_numRows + Repeat1__numRows
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
      <p><font face="Arial, Helvetica, sans-serif"><b>ENVIAR</b></font> 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif"> </font> 
    </td>
    <td> 
      <form name="enviar" method="post" action="<%=MM_editAction%>" enctype="multipart/form-data">
        <p> 
          <input type="hidden" name="id" value="<%=(rsrecomendados.Fields.Item("ID").Value)%>">
          <b>Nombre:</b> <%=(rsrecomendados.Fields.Item("nombre").Value)%><br>
          <b>Email: </b>&nbsp;&nbsp;&nbsp; <%=(rsrecomendados.Fields.Item("email").Value)%></p>
        <p><b>Contestacion:</b><br>
          <%=(rsrecomendados.Fields.Item("contestacion").Value)%><br>
        </p>
        <p><b>Productos:</b><br>
          <%=(rsrecomendados.Fields.Item("productos").Value)%> 
        </p>
        <p><b>Total:</b> <%=(rsrecomendados.Fields.Item("total").Value)%></p>
        <p> 
          <input type="hidden" name="enviados" value="si">
          <input type="submit" name="enviar" value="enviar">
        </p>
        <input type="hidden" name="MM_recordId" value="<%= rsrecomendados.Fields.Item("ID").Value %>">
        <input type="hidden" name="MM_update" value="true">
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
