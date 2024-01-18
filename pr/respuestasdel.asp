<%@LANGUAGE="VBSCRIPT"%><%
' *** Delete Record: construct a sql delete statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_tableName = "recomendados"
  MM_tableCol = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_redirectPage = "respuestasviewalladm.asp"

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
<title>Borrar respuesta</title>
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
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>ID</b>: <%=(rsrecomendados.Fields.Item("ID").Value)%><br>
          <b>Fecha: </b><%=(rsrecomendados.Fields.Item("fecha").Value)%></font></p>
        <p><b><font face="Verdana, Arial, Helvetica, sans-serif">Nombre : </font></b><font face="Verdana, Arial, Helvetica, sans-serif"><%=(rsrecomendados.Fields.Item("nombre").Value)%><br>
          </font><font face="Verdana, Arial, Helvetica, sans-serif"><b>Apellido:</b> 
          <%=(rsrecomendados.Fields.Item("apellido").Value)%></font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"> <b>Email:</b> <%=(rsrecomendados.Fields.Item("email").Value)%><br>
          </font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Contestacion:<br>
          </b><%=(rsrecomendados.Fields.Item("contestacion").Value)%></font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Productos:</b><br>
          <%=(rsrecomendados.Fields.Item("productos").Value)%><br>
          </font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Total:</b> <%=(rsrecomendados.Fields.Item("total").Value)%></font></p>
        <p> <font face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="submit" name="borrar" value="BORRAR">
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="respuestasviewalladm.asp">cancelar</a> 
          </font> </p>
        <input type="hidden" name="MM_recordId" value="<%= rsrecomendados.Fields.Item("ID").Value %>">
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
