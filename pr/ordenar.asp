<%@LANGUAGE="VBSCRIPT"%>
<%

Dim rscheckduplicates__strID
rscheckduplicates__strID = "0"
if(Request.QueryString("ID") <> "") then rscheckduplicates__strID = Request.QueryString("ID")

%>
<%
set rscheckduplicates = Server.CreateObject("ADODB.Recordset")
rscheckduplicates.ActiveConnection = "dsn=5084.quotes;"
rscheckduplicates.Source = "SELECT *  FROM ordenes  WHERE ID = " + Replace(rscheckduplicates__strID, "'", "''") + ""
rscheckduplicates.CursorType = 0
rscheckduplicates.CursorLocation = 2
rscheckduplicates.LockType = 3
rscheckduplicates.Open
rscheckduplicates_numRows = 0
%> 
<% If not rscheckduplicates.eof then
response.redirect ("duplicate.asp")
else%> <%
' *** Insert Record: construct a sql insert statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
If (CStr(Request("MM_insert")) <> "") Then

  MM_tableName = "ordenes"
  MM_fields = "id,ID,none,none,NULL,productos,productos,',none,'',total,total,none,none,NULL,fecha,fecha,#,none,NULL,aprobado,aprobado,',none,'',name,nombre,',none,'',apellido,apellido,',none,'',email,email,',none,'',direccion,direccion,',none,'',ciudad,ciudad,',none,'',estado,estado,',none,'',codigo,codigo,',none,'',telefono,telefono,',none,'',tipo,tarjeta,',none,'',numero,numero,',none,'',expiracion,expiracion,',none,'',comentarios,comentarios,',none,''"
  MM_redirectPage = "ordenar1.asp"

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
set rsordenes = Server.CreateObject("ADODB.Recordset")
rsordenes.ActiveConnection = "dsn=5084.quotes;"
rsordenes.Source = "SELECT * FROM ordenes"
rsordenes.CursorType = 0
rsordenes.CursorLocation = 2
rsordenes.LockType = 3
rsordenes.Open
rsordenes_numRows = 0
%> 
<html><!-- #BeginTemplate "/Templates/basic.dwt" --><!-- DW6 -->

<head>
<!-- #BeginEditable "doctitle" --> 
<title>Untitled Document</title>
<!-- #EndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" onLoad="MM_preloadImages('../images/recetascaseras1.gif','../images/homefoto.gif','../images/home1.gif','../images/enfermedadesfoto.gif','../images/enfermedades1.gif','../images/topicosfoto.gif','../images/topicos1.gif','../images/preguntasfoto.gif','../images/preguntas1.gif','../images/drarmozafoto.gif','../images/drarmoza1.gif','../images/recetascaserasfoto.gif')">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td align="right" width="10%" valign="bottom" background="../images/square.gif"><img src="../images/homefoto.gif" width="150" height="90" name="foto"></td>
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center" background="../images/square.gif"><img src="../images/logo.gif" width="377" height="90"></td>
          <td align="right" background="../images/square.gif"><img src="../images/right.gif" width="13" height="90"></td>
          <td width="10%" bgcolor="#006FA4">
            <table border="0" cellpadding="0" cellspacing="0" align="center">
              <tr align="center"> 
                <td><img src="../images/buscar.gif"></td>
              </tr>
              <tr align="center"> 
                <td> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr valign="bottom"> 
                      <td> 
                        <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.nombre.value == "")
  {
    alert("Registration Confirmation");
    theForm.nombre.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="/asp/enfnom.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
                          <!--webbot bot="Validation" B-Value-Required="TRUE" --> 
                          <input type="text" name="nombre" size="13">
                          <input type="image" img src="/images/ir.gif" name="buscar" border="0" width="25" height="19">
                        </form>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="10%" height="10" bgcolor="#006FA4"> <img src="../images/barra.gif" width="133" height="15"> 
    </td>
    <td valign="bottom" bgcolor="#006FA4" height="10">&nbsp; </td>
  </tr>
  <tr> 
    <td valign="top" align="right" width="10%" bgcolor="#006FA4"> 
      <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../images/top.gif" width="150"></td>
        </tr>
        <tr> 
          <td><a href="../index.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('foto','','../images/homefoto.gif','home','','../images/home1.gif',1)"><img src="../images/home.gif" width="150" height="30" name="home" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="../asp/enfermedades.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('foto','','../images/enfermedadesfoto.gif','enfermedades','','../images/enfermedades1.gif',1)"><img src="../images/enfermedades.gif" width="150" height="30" name="enfermedades" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="../asp/articulos.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('foto','','../images/topicosfoto.gif','topicos','','../images/topicos1.gif',1)"><img src="../images/topicos.gif" width="150" height="30" name="topicos" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="../asp/preguntas.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('foto','','../images/preguntasfoto.gif','preguntas','','../images/preguntas1.gif',1)"><img src="../images/preguntas.gif" width="150" height="30" name="preguntas" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="../asp/recetascaseras.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('foto','','../images/recetascaserasfoto.gif','recetascaseras','','../images/recetascaseras1.gif',1)"><img src="../images/recetascaseras.gif" name="recetascaseras" width="150" height="30" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="../drarmoza.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('foto','','../images/drarmozafoto.gif','drarmoza','','../images/drarmoza1.gif',1)"><img src="../images/drarmoza.gif" width="150" height="30" name="drarmoza" border="0"></a></td>
        </tr>
        <tr> 
          <td><img src="../images/bottom.gif" width="150"></td>
        </tr>
        <tr bgcolor="#006FA4" align="center" valign="middle"> 
          <td><a href="../preguntele.htm"><img src="../images/preguntele.gif" width="84" height="84" border="0"></a></td>
        </tr>
      </table>
    </td>
    <td align="left" valign="top"><!-- #BeginEditable "contenido" --> 
      <table border="0" width="100%">
        <tr> 
          <td width="20"></td>
          <td height="5">&nbsp; </td>
        </tr>
        <tr> 
          <td width="20"></td>
          <td height="30"> 
            <form name="form1" method="post" action="<%=MM_editAction%>" onSubmit="MM_validateForm('name','','R','apellido','','R','email','','NisEmail','direccion','','R','ciudad','','R','estado','','R','codigo','','R','telefono','','R');return document.MM_returnValue">
              <p> 
                <input type="hidden" name="id" value="<%=(rsrecomendados.Fields.Item("ID").Value)%>">
                <input type="hidden" name="productos" value="<%=(rsrecomendados.Fields.Item("productos").Value)%>">
                <input type="hidden" name="total" value="<%=(rsrecomendados.Fields.Item("total").Value)%>">
                <font face="Verdana, Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="fecha" value="<%= date() %>">
                <input type="hidden" name="aprobado" value="no">
                </font> </p>
              <p>&nbsp; </p>
              <pre>Nombre:    <input type="text" name="name" size="20">  Apellido:  <input type="text" name="apellido" size="20">
Email:     <input type="text" name="email" size="50">
Dirección: <input type="text" name="direccion" size="50">
Ciudad:    <input type="text" name="ciudad" size="20">   Estado: <input type="text" name="estado" size="5">  Código: <input type="text" name="codigo" size="10">
Teléfono:  <input type="text" name="telefono" size="20">

 
Forma de pago:       <select name="tipo" size="1"><option value="western">Western Union</option><option value="visa">Visa</option><option value="mastercard">Mastercard</option><option value="american">American Express</option><option value="discover">Discover</option></select>
Numero de Tarjeta:   <input type="text" name="numero" size="30">
Fecha de expiración: <input type="text" name="expiracion" size="20" onBlur="MM_validateForm('fecha','','R');return document.MM_returnValue">



Comentarios:
<textarea name="comentarios" cols="50"></textarea>

<input type="submit" name="mandar" value="Ordenar"></pre>
              <input type="hidden" name="MM_insert" value="true">
              <%Session("svid")= (rsrecomendados.Fields.Item("id").Value) %> 
            </form>
            <p>&nbsp; 
          </td>
        </tr>
        <tr> 
          <td width="20" align="right"></td>
          <td></td>
        </tr>
        <tr> 
          <td width="20" align="right"></td>
          <td height="5"></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <!-- #EndEditable --></td>
  </tr>
  <tr valign="middle"> 
    <td width="10%" bgcolor="#89B0D8" height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="1"><a href="/disclaimer.htm">DISCLAIMER</a></font></td>
    <td bgcolor="#89B0D8" align="center" valign="middle"><font face="Arial, Helvetica, sans-serif" size="1">[<a href="../index.htm">HOME</a>] 
      [<a href="../asp/enfermedades.asp">ENFERMEDADES</a>] [<a href="../asp/articulos.asp">TOPICOS</a>] 
      [<a href="../asp/articulos.asp">PREGUNTAS</a>] [<a href="../asp/recetascaseras.asp">RECETAS</a>] 
      [<a href="../drarmoza.htm">C.ARMOZA</a>]</font></td>
  </tr>
  <tr>
    <td valign="top" align="right" width="10%">&nbsp;</td>
    <td align="center" valign="middle"><font size="1">Copyright © 2000-2003 NuestraMedicina.com. 
      All rights reserved.</font><font face="Arial, Helvetica, sans-serif" size="1">&nbsp;</font></td>
  </tr>
</table>
</body>
<!-- #EndTemplate --></html>
