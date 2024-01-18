<%@LANGUAGE="VBSCRIPT"%><%
' *** Update Record: construct a sql update statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_tableName = "cuestionarios"
  MM_tableCol = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_fields = "respondido,respondido,',none,''"
  MM_redirectPage = "cuestionarioall.asp"

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

Dim rscuestionario__strID
rscuestionario__strID = "0"
if(Session("svID") <> "") then rscuestionario__strID = Session("svID")

%> <%
set rscuestionario = Server.CreateObject("ADODB.Recordset")
rscuestionario.ActiveConnection = "dsn=5084.quotes;"
rscuestionario.Source = "SELECT *  FROM cuestionarios  WHERE ID = " + Replace(rscuestionario__strID, "'", "''") + ""
rscuestionario.CursorType = 0
rscuestionario.CursorLocation = 2
rscuestionario.LockType = 3
rscuestionario.Open
rscuestionario_numRows = 0
 %> 
<html>
<head>
<title>Confirmacion de respuesta</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF">
<table width="75%" border="0" cellspacing="10" cellpadding="10">
  <tr align="center"> 
    <td> 
      <p><%
id =(rscuestionario.Fields.Item("ID").Value)
email = (rscuestionario.Fields.Item("email").Value)


Set objMail = CreateObject("CDONTS.Newmail")
objMail.From ="drarmoza@hotmail.com"
objMail.To = email
objMail.Subject = "Respuesta al cuestionario"
objMail.Body = "Esta es su contestacion del Dr. Cesar Armoza, para ver la respuesta a su cuestionario haga click aqui     http://www.nuestramedicina.com/pr/detalle.asp?id=" & id
objMail.Send
  %><font face="Arial, Helvetica, sans-serif"><b><font face="Verdana, Arial, Helvetica, sans-serif"><br>
        Su respuesta ha sido enviada</font></b></font></p>
      <form name="form1" method="post" action="<%=MM_editAction%>">
        <input type="hidden" name="respondido" value="si">
        <input type="submit" name="Submit" value="Volver y actualizar cuestionario">
        <input type="hidden" name="MM_recordId" value="<%= rscuestionario.Fields.Item("ID").Value %>">
        <input type="hidden" name="MM_update" value="true">
      </form>
      <p><font face="Arial, Helvetica, sans-serif"><b></b></font></p>
      </td>
  </tr>
</table>
</body>
</html>
