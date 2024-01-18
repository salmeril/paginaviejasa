<%@LANGUAGE="VBSCRIPT"%> <%
set rsrecomendados = Server.CreateObject("ADODB.Recordset")
rsrecomendados.ActiveConnection = "dsn=5084.quotes;"
rsrecomendados.Source = "SELECT * FROM recomendados ORDER BY ID DESC"
rsrecomendados.CursorType = 0
rsrecomendados.CursorLocation = 2
rsrecomendados.LockType = 3
rsrecomendados.Open
rsrecomendados_numRows = 0
%><%
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
    <td width="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <td height="5" align="right"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330066">PRODUCTOS 
        RECOMENDADOS</font><font color="#3366CC"> </font></b></font></div>
    </td>
  </tr>
  <tr> 
    <td width="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <td height="30"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif"><a href="add.asp">Agregar</a></font> 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif"><font face="Verdana, Arial, Helvetica, sans-serif"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="2"></font></font></font> 
        </font> 
    </td>
    <td> 
      <table width="100%" border="1" cellspacing="0" cellpadding="0">
        <tr bgcolor="#3399CC"> 
          <td width="30"><font face="Verdana, Arial, Helvetica, sans-serif"><b>ID</b></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif"><b>NOMBRE</b></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif"><b>EMAIL</b></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
        </tr>
        <%
While ((Repeat1__numRows <> 0) AND (NOT rsrecomendados.EOF))
%> 
        <tr> 
          <td width="30"><font face="Verdana, Arial, Helvetica, sans-serif"><%=(rsrecomendados.Fields.Item("ID").Value)%></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif"><%=(rsrecomendados.Fields.Item("nombre").Value)%></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif"><%=(rsrecomendados.Fields.Item("email").Value)%> 
            </font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif"><a HREF="edit.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & rsrecomendados.Fields.Item("ID").Value %>">Editar</a></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif"><a href="delete.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & rsrecomendados.Fields.Item("ID").Value %>">Borrar</a></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif"><A HREF="enviar.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & rsrecomendados.Fields.Item("ID").Value %>">Enviar</A></font></td>
        </tr>
        <%
Repeat1__index=Repeat1__index+1
Repeat1__numRows=Repeat1__numRows-1
rsrecomendados.MoveNext()
Wend
%> 
      </table>
      <p>&nbsp; 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <td height="5"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
  </tr>
</table>
<p>&nbsp; </p>
</body>
</html>
