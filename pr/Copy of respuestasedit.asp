<%@LANGUAGE="VBSCRIPT"%> <%
' *** Update Record: construct a sql update statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_tableName = "recomendados"
  MM_tableCol = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_fields = "nombre,nombre,',none,'',apellido,apellido,',none,'',email,email,',none,'',contestacion,contestacion,',none,'',productos,productos,',none,'',total,total,none,none,NULL,fechaenvio,fechaenvio,#,none,NULL,enviado,enviado,',none,''"
  MM_redirectPage = "enviar1.asp"

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
rsrecomendados.Source = "SELECT *  FROM recomendados  WHERE ID = " + Replace(rsrecomendados__MMColParam, "'", "''") + ""
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
<title>Untitled Document</title>
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
<p>&nbsp; </p>
<table border="0" width="100%">
  <tr> 
    <td width="20">&nbsp;</td>
    <td height="5">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20">&nbsp;</td>
    <td height="30" bgcolor="#FFFFFF"> 
      <p>&nbsp; 
    </td>
  </tr>
  <tr> 
    <td width="20" align="right"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif"> </font> 
    </td>
    <td> 
      <form name="UPDATE" method="post" action="<%=MM_editAction%>">
        <p><%=(rsrecomendados.Fields.Item("ID").Value)%></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Nombre:</b> 
          <input type="text" name="nombre" value="<%=(rsrecomendados.Fields.Item("nombre").Value)%>">
          <br>
          <b>Apellido:</b> </font><font face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="text" name="apellido" value="<%=(rsrecomendados.Fields.Item("apellido").Value)%>">
          <br>
          <b>Email:</b> &nbsp;&nbsp;&nbsp; 
          <input type="text" name="email" value="<%=(rsrecomendados.Fields.Item("email").Value)%>">
          <br>
          </font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Contestacion:<br>
          <textarea name="contestacion" cols="50" rows="5"><%=(rsrecomendados.Fields.Item("contestacion").Value)%></textarea>
          </b></font></p>
        <p><b><font face="Verdana, Arial, Helvetica, sans-serif">Productos</font></b><font face="Verdana, Arial, Helvetica, sans-serif">:<br>
          <textarea name="productos" cols="50" rows="5"><%=(rsrecomendados.Fields.Item("productos").Value)%></textarea>
          </font></p>
        <p><font face="Verdana, Arial, Helvetica, sans-serif"><b>Total</b>: 
          <input type="text" name="total" value="<%=(rsrecomendados.Fields.Item("total").Value)%>">
          <input type="hidden" name="fechaenvio" value="<%= date() %>">
          <input type="hidden" name="enviado" value="si">
          </font></p>
        <p> 
          <input type="submit" name="enviar" value="ENVIAR" onClick="MM_validateForm('nombre','','R','email','','RisEmail','contestacion','','R');return document.MM_returnValue">
          <font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="respuestasview.asp">cancelar</a> 
          </font> </p>
        <input type="hidden" name="MM_recordId" value="<%= rsrecomendados.Fields.Item("ID").Value %>">
        <input type="hidden" name="MM_update" value="true">
        <%Session("svid")= (rsrecomendados.Fields.Item("ID").Value) %> 
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






