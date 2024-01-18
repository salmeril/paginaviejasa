<%@LANGUAGE="VBSCRIPT"%> <%
set rsordenes = Server.CreateObject("ADODB.Recordset")
rsordenes.ActiveConnection = "dsn=5084.quotes;"
rsordenes.Source = "SELECT * FROM ordenes ORDER BY ID DESC"
rsordenes.CursorType = 0
rsordenes.CursorLocation = 2
rsordenes.LockType = 3
rsordenes.Open
rsordenes_numRows = 0
%><%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsordenes_numRows = rsordenes_numRows + Repeat1__numRows
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
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#33CCCC">ORDENES</font></b></font></div>
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
        <tr bgcolor="#33CCCC"> 
          <td width="30" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>ID</b></font></td>
          <td height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>NOMBRE</b></font></td>
          <td height="23" bgcolor="#33CCCC"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>EMAIL</b></font></td>
          <td height="23"><b>Aprobado</b></td>
          <td height="23"><b>Fecha</b></td>
          <td height="23">.</td>
        </tr>
        <%
While ((Repeat1__numRows <> 0) AND (NOT rsordenes.EOF))
%> 
        <tr> 
          <td width="30"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rsordenes.Fields.Item("ID").Value)%></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rsordenes.Fields.Item("nombre").Value)%>&nbsp;&nbsp;<%=(rsordenes.Fields.Item("apellido").Value)%></font></td>
          <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rsordenes.Fields.Item("email").Value)%> 
            </font></td>
          <td><%=(rsordenes.Fields.Item("aprobado").Value)%></td>
          <td><%=(rsordenes.Fields.Item("fechaaprobado").Value)%></td>
          <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><A HREF="aprobar.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & rsordenes.Fields.Item("ID").Value %>">Aprobar</A></font></td>
        </tr>
        <%
Repeat1__index=Repeat1__index+1
Repeat1__numRows=Repeat1__numRows-1
rsordenes.MoveNext()
Wend
%> 
      </table>
      <p>&nbsp; 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    <td height="5"><a href="ordenview.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">ordenes 
      no aprobadas</font></a></td>
  </tr>
</table>
<p>&nbsp; </p>
</body>
</html>

