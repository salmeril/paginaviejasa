<%@LANGUAGE="VBSCRIPT"%> <%

Dim rscuestionarios__MMColParam
rscuestionarios__MMColParam = "0"
if(Request.Form("fechainicial") <> "") then rscuestionarios__MMColParam = Request.Form("fechainicial")

%> <%
set rscuestionarios = Server.CreateObject("ADODB.Recordset")
rscuestionarios.ActiveConnection = "dsn=5084.quotes;"
rscuestionarios.Source = "SELECT *  FROM cuestionarios  WHERE fecha >=#" + Replace(rscuestionarios__MMColParam, "'", "''") + "# and lugar='USA'"
rscuestionarios.CursorType = 0
rscuestionarios.CursorLocation = 2
rscuestionarios.LockType = 3
rscuestionarios.Open
rscuestionarios_numRows = 0
%> 
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rscuestionarios_numRows = rscuestionarios_numRows + Repeat1__numRows
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
<title>Listado de emails y telefonos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF">
<p>&nbsp;</p>
<table border="0" width="100%">
  <tr> 
    <td width="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <td height="5" align="right"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#000000">LISTADO 
        DE EMAIL Y TELEFONOS</font></b></font></div>
    </td>
  </tr>
  <tr> 
    <td width="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <td height="30"> 
      <p>&nbsp; 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif"><font face="Verdana, Arial, Helvetica, sans-serif"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="2"></font></font></font> 
        </font> 
    </td>
    <td> 
      <table width="100%" border="1" cellspacing="0" cellpadding="0">
        <tr bgcolor="#CCCCCC"> 
          <td width="30"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Fecha</font></b></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Nombre</b></font></td>
          <td height="23"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Estado</font></b></td>
          <td height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Email</b></font></td>
          <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Telefono</font></b></td>
          <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Horario</font></b></td>
        </tr>
        <%
While ((Repeat1__numRows <> 0) AND (NOT rscuestionarios.EOF))
%> 
        <tr> 
          <td width="30"><%=(rscuestionarios.Fields.Item("fecha").Value)%></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
            &nbsp; <%=(rscuestionarios.Fields.Item("nombre").Value)%> 
            <%=(rscuestionarios.Fields.Item("apellido").Value)%> 
            </font></td>
          <td><%=(rscuestionarios.Fields.Item("estado").Value)%></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rscuestionarios.Fields.Item("email").Value)%> 
            </font></td>
          <td><%=(rscuestionarios.Fields.Item("telefono").Value)%></td>
          <td><%=(rscuestionarios.Fields.Item("horario").Value)%></td>
        </tr>
        <%
Repeat1__index=Repeat1__index+1
Repeat1__numRows=Repeat1__numRows-1
rscuestionarios.MoveNext()
Wend
%> 
      </table>
      <p>&nbsp; 
    </td>
  </tr>
</table>
<p>&nbsp; </p>
</body>
</html>

