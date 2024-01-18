<%@LANGUAGE="VBSCRIPT"%> <%

Dim rsrecomendados__strID
rsrecomendados__strID = "0"
if(Session("svID") <> "") then rsrecomendados__strID = Session("svID")

%> <%
set rsrecomendados = Server.CreateObject("ADODB.Recordset")
rsrecomendados.ActiveConnection = "dsn=5084.quotes;"
rsrecomendados.Source = "SELECT *  FROM recomendados  WHERE ID = " + Replace(rsrecomendados__strID, "'", "''") + ""
rsrecomendados.CursorType = 0
rsrecomendados.CursorLocation = 2
rsrecomendados.LockType = 3
rsrecomendados.Open
rsrecomendados_numRows = 0
%> 
<%
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
<title>respuesta duplicada</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td align="left" valign="top"> 
      <table width="75%" border="0" cellspacing="10" cellpadding="10">
        <tr align="center"> 
          <td> 
            <p>&nbsp; </p>
            <p><font face="Arial, Helvetica, sans-serif"><font face="Verdana, Arial, Helvetica, sans-serif"><b>Su 
              respuesta ya fue enviada previamente</b></font></font></p>
            <p>para modificar la respuesta y enviarla nuevamente haga <a HREF="respuestasedit.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & rsrecomendados.Fields.Item("ID").Value %>">click 
              aqui</a></p>
            <p>&nbsp;</p>
            <p><font face="Arial, Helvetica, sans-serif"><b></b></font></p>
            <p>&nbsp;</p>
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
</body>
</html>

