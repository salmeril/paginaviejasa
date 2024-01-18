<%@LANGUAGE="VBSCRIPT"%> <%

Dim rscheckduplicates__strID
rscheckduplicates__strID = "0"
if(session("svID") <> "") then rscheckduplicates__strID = session("svID")

%> <%
set rscheckduplicates = Server.CreateObject("ADODB.Recordset")
rscheckduplicates.ActiveConnection = "dsn=5084.quotes;"
rscheckduplicates.Source = "SELECT *  FROM recomendados  WHERE ID = " + Replace(rscheckduplicates__strID, "'", "''") + ""
rscheckduplicates.CursorType = 0
rscheckduplicates.CursorLocation = 2
rscheckduplicates.LockType = 3
rscheckduplicates.Open
rscheckduplicates_numRows = 0
%> 
<% If not rscheckduplicates.eof then
response.redirect ("respduplicate.asp")
else%><%
' *** Insert Record: construct a sql insert statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
If (CStr(Request("MM_insert")) <> "") Then

  MM_tableName = "recomendados"
  MM_fields = "fecha,fecha,#,none,NULL,enviado,enviado,',none,'',id,ID,none,none,NULL,nombre,nombre,',none,'',apellido,apellido,',none,'',email,email,',none,'',contestacion,contestacion,',none,'',productos,productos,',none,'',total,total,none,none,NULL"
  MM_redirectPage = "view.asp"

  ' create the insert sql statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_fieldsArray(i+1)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_insertStr = "insert into " & MM_tableName & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  ' finish the sql and execute it
  Set MM_insertCmd = Server.CreateObject("ADODB.Command")
  MM_insertCmd.ActiveConnection = "dsn=5084.quotes;"
  MM_insertCmd.CommandText = MM_insertStr
  MM_insertCmd.Execute

  ' redirect with URL parameters
  If (MM_redirectPage = "") Then
    MM_redirectPage = CStr(Request("URL"))
  End If
  If (InStr(1, MM_redirectPage, "?", vbTextCompare) = 0 And (Request.QueryString <> "")) Then
    MM_redirectPage = MM_redirectPage & "?" & Request.QueryString
  End If
  Call Response.Redirect(MM_redirectPage)
End If
%> 
<% End If %> 
<%
set rsrecomendados = Server.CreateObject("ADODB.Recordset")
rsrecomendados.ActiveConnection = "dsn=5084.quotes;"
rsrecomendados.Source = "SELECT * FROM recomendados"
rsrecomendados.CursorType = 0
rsrecomendados.CursorLocation = 2
rsrecomendados.LockType = 3
rsrecomendados.Open
rsrecomendados_numRows = 0
%> 
<html>
<head>
<title>Responder Cuestionario</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_validateForm() { //v3.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (val!=''+num) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF">
<table border="0" width="100%">
  <tr> 
    <td width="20">&nbsp;</td>
    <td height="5"><font face="Verdana, Arial, Helvetica, sans-serif"><b>AGREGAR 
      CONTESTACION </b></font></td>
  </tr>
  <tr> 
    <td width="20">&nbsp;</td>
    <td height="30"> 
      <p>&nbsp; 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif"> </font> 
    </td>
    <td> 
      <form name="productos recomendados" method="post" action="<%=MM_editAction%>">
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b> 
          <input type="hidden" name="fecha" value="<%= date() %>">
          <input type="hidden" name="enviado" value="no">
          <input type="hidden" name="id" value="<%=session("svid")%>">
          </b></font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Nombre:</b> 
          <input type="text" name="nombre" value="<%=session("svnombre")%>">
          <br>
          <b>Apellido:</b> 
          <input type="text" name="apellido" value="<%=session("svapellido")%>">
          <br>
          <b>Email:</b> &nbsp;&nbsp;&nbsp; 
          <input type="text" name="email" value="<%=session("svemail")%>">
          <br>
          </font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Contestacion:<br>
          <textarea name="contestacion" cols="50" rows="5"></textarea>
          </b></font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Productos:</b><br>
          <textarea name="productos" cols="50" rows="5"></textarea>
          </font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Total:</b> 
          <input type="text" name="total">
          </font></p>
        <p> 
          <input type="submit" name="agregar" value="AGREGAR" onClick="MM_validateForm('nombre','','R','email','','RisEmail','contestacion','','R');return document.MM_returnValue">
        </p>
        <input type="hidden" name="MM_insert" value="true">
      </form>
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> </font> 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right">&nbsp;</td>
    <td height="5">&nbsp;</td>
  </tr>
</table>
</body>
</html>

