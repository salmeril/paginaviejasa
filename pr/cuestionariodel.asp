<%@LANGUAGE="VBSCRIPT"%><%
' *** Delete Record: construct a sql delete statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_tableName = "cuestionarios"
  MM_tableCol = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_redirectPage = "cuestionarioall.asp"

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

Dim rscuestionario__MMColParam
rscuestionario__MMColParam = "0"
if(Request.QueryString("ID") <> "") then rscuestionario__MMColParam = Request.QueryString("ID")

%> <%
set rscuestionario = Server.CreateObject("ADODB.Recordset")
rscuestionario.ActiveConnection = "dsn=5084.quotes;"
rscuestionario.Source = "SELECT * FROM cuestionarios WHERE ID = " + Replace(rscuestionario__MMColParam, "'", "''") + ""
rscuestionario.CursorType = 0
rscuestionario.CursorLocation = 2
rscuestionario.LockType = 3
rscuestionario.Open
rscuestionario_numRows = 0
%> 
<html><!-- #BeginTemplate "/Templates/basic.dwt" --><!-- DW6 -->

<head>
<!-- #BeginEditable "doctitle" --> 
<title>Cuestionario delete</title>
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
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="15">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td width="15" height="28">&nbsp;</td>
          <td height="28"><b><font color="#333399">CUESTIONARIO DE SALUD</font></b></td>
        </tr>
        <tr> 
          <td width="15">&nbsp;</td>
          <td> 
            <form name="form1" method="post" action="<%=MM_editAction%>">
              <pre>

ID: <b><%=(rscuestionario.Fields.Item("ID").Value)%></b>
Fecha: <b><%=(rscuestionario.Fields.Item("fecha").Value)%></b>     

Nombre: <b><%=(rscuestionario.Fields.Item("nombre").Value)%></b><br>Apellido: <b><%=(rscuestionario.Fields.Item("apellido").Value)%></b>
<br>Edad: <b><%=(rscuestionario.Fields.Item("edad").Value)%></b>
<br>Dirección: <b><%=(rscuestionario.Fields.Item("direccion").Value)%></b><br>Ciudad: <b><%=(rscuestionario.Fields.Item("ciudad").Value)%></b>
Estado: <b><%=(rscuestionario.Fields.Item("estado").Value)%></b>    Código: <b><%=(rscuestionario.Fields.Item("codigo").Value)%></b>
<br>Teléfono: <b><%=(rscuestionario.Fields.Item("telefono").Value)%></b><br>Email:    <b><%=(rscuestionario.Fields.Item("email").Value)%></b>

Como desea ser contactado? <b><%=(rscuestionario.Fields.Item("contacto").Value)%></b><br>Horario conveniente para llamarlo: <b><%=(rscuestionario.Fields.Item("horario").Value)%></b>

Problema principal y pregunta? <br><b><%=(rscuestionario.Fields.Item("pregunta").Value)%></b>

Que le dijo su doctor acerca de su salud? <b><%=(rscuestionario.Fields.Item("salud").Value)%></b>
Esta recibiendo tratamiento medico?   <b><%=(rscuestionario.Fields.Item("tratamiento").Value)%></b>
Desde cuando? <b><%=(rscuestionario.Fields.Item("desde").Value)%></b> 
Que medicinas esta tomando? <b><%=(rscuestionario.Fields.Item("medicinas").Value)%></b>
Que vitaminas esta tomando? <b><%=(rscuestionario.Fields.Item("vitaminas").Value)%></b>

Es usted alérgico a algun medicamento? <b><%=(rscuestionario.Fields.Item("alergico").Value)%></b>
Cual? <b><%=(rscuestionario.Fields.Item("alergias").Value)%></b>

Alguna infección?  <b><%=(rscuestionario.Fields.Item("infeccion").Value)%></b>
Cual? <b><%=(rscuestionario.Fields.Item("infecciones").Value)%></b>

Ha sido operado? <b><%=(rscuestionario.Fields.Item("operado").Value)%></b>
De que? <b><%=(rscuestionario.Fields.Item("operaciones").Value)%></b>

Tiene o ha tenido:

<input type="checkbox" name="hiv" <%If (rscuestionario.Fields.Item("hiv").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>HIV   <input type="checkbox" name="aid" value="yes" <%If (rscuestionario.Fields.Item("aid").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>AID   <input type="checkbox" name="hepatitis" value="yes" <%If (rscuestionario.Fields.Item("hepatitis").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>HEPATITIS    <input type="checkbox" name="venereas" value="yes" <%If (rscuestionario.Fields.Item("venereas").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ENFERMEDADES VENEREAS
<input type="checkbox" name="vista" value <%If (rscuestionario.Fields.Item("vista").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PROBLEMAS DE VISTA    <input type="checkbox" name="memoria" value="yes" <%If (rscuestionario.Fields.Item("memoria").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>MEMORIA        <input type="checkbox" name="insomnio" value="yes" <%If (rscuestionario.Fields.Item("insomnio").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>INSOMNIO
<input type="checkbox" name="oido" value="yes" <%If (rscuestionario.Fields.Item("oido").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PROBLEMAS DE OIDO     <input type="checkbox" name="colesterol" value="yes" <%If (rscuestionario.Fields.Item("colesterol").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>COLESTEROL     <input type="checkbox" name="diabetes" value="yes" <%If (rscuestionario.Fields.Item("diabetes").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DIABETES
<input type="checkbox" name="nariz" value="yes" <%If (rscuestionario.Fields.Item("nariz").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ALERGIAS NARIZ Y PIEL <input type="checkbox" name="nervios" value="yes" <%If (rscuestionario.Fields.Item("nervios").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>NERVIOS        <input type="checkbox" name="anemia" value="yes" <%If (rscuestionario.Fields.Item("anemia").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ANEMIA
<input type="checkbox" name="bronquitis" value="yes" <%If (rscuestionario.Fields.Item("bronquitis").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>BRONQUITIS            <input type="checkbox" name="depresion" value="yes" <%If (rscuestionario.Fields.Item("depresion").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DEPRESION      <input type="checkbox" name="asma" value="yes" <%If (rscuestionario.Fields.Item("asma").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ASMA
<input type="checkbox" name="catarro" value="yes" <%If (rscuestionario.Fields.Item("catarro").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>MUCHO CATARRO         <input type="checkbox" name="ansiedad" value="yes" <%If (rscuestionario.Fields.Item("ansiedad").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ANSIEDAD       <input type="checkbox" name="alcoholismo" value="yes" <%If (rscuestionario.Fields.Item("alcoholismo").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ALCOHOLISMO
<input type="checkbox" name="acidez" value="yes" <%If (rscuestionario.Fields.Item("acidez").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ACIDEZ ESTOMACAL      <input type="checkbox" name="cansancio" value="yes" <%If (rscuestionario.Fields.Item("cansancio").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>CANSANCIO      <input type="checkbox" name="drogadiccion" value="yes" <%If (rscuestionario.Fields.Item("drogadiccion").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DROGADICCION
<input type="checkbox" name="estomago" value="yes" <%If (rscuestionario.Fields.Item("estomago").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DOLOR DE ESTOMAGO     <input type="checkbox" name="migrana" value="yes" <%If (rscuestionario.Fields.Item("migrana").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>MIGRAÑA        <input type="checkbox" name="prostata" value="yes" <%If (rscuestionario.Fields.Item("prostata").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PROSTATA
<input type="checkbox" name="cabeza" value="yes" <%If (rscuestionario.Fields.Item("cabeza").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DOLORES DE CABEZA     <input type="checkbox" name="gases" value="yes" <%If (rscuestionario.Fields.Item("gases").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>GASES          <input type="checkbox" name="ovarios" value="yes" <%If (rscuestionario.Fields.Item("ovarios").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>OVARIOS
<input type="checkbox" name="apetito" value="yes" <%If (rscuestionario.Fields.Item("apetito").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>APETITO MUCHO POCO    <input type="checkbox" name="rinones" value="yes" <%If (rscuestionario.Fields.Item("rinones").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>RIÑONES        <input type="checkbox" name="vaginal" value="yes" <%If (rscuestionario.Fields.Item("vaginal").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>INFECCION VAGINAL
<input type="checkbox" name="menstruales" value="yes" <%If (rscuestionario.Fields.Item("menstruales").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PROBLEMAS MENSTRUALES <input type="checkbox" name="orina" value="yes" <%If (rscuestionario.Fields.Item("orina").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ORINA MUCHO    <input type="checkbox" name="diarrea" value="yes" <%If (rscuestionario.Fields.Item("diarrea").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DIARREA
<input type="checkbox" name="orinar" value="yes" <%If (rscuestionario.Fields.Item("orinar").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>MOLESTIAS AL ORINAR   <input type="checkbox" name="estrenimiento" value="yes" <%If (rscuestionario.Fields.Item("estrenimiento").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ESTREÑIMIENTO  <input type="checkbox" name="embarazo" value="yes" <%If (rscuestionario.Fields.Item("embarazo").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>EMBARAZO
<input type="checkbox" name="encias" value="yes" <%If (rscuestionario.Fields.Item("encias").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ENCIAS IMFLAMADAS     <input type="checkbox" name="reumatismo" value="yes" <%If (rscuestionario.Fields.Item("reumatismo").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>REUMATISMO     <input type="checkbox" name="impotencia" value="yes" <%If (rscuestionario.Fields.Item("impotencia").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>IMPOTENCIA
<input type="checkbox" name="higado" value="yes" <%If (rscuestionario.Fields.Item("higado").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>HIGADO                <input type="checkbox" name="artritis" value="yes" <%If (rscuestionario.Fields.Item("artritis").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>ARTRITIS       <input type="checkbox" name="frigidez" value="yes" <%If (rscuestionario.Fields.Item("frigidez").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>FRIGIDEZ
<input type="checkbox" name="piernas" value="yes" <%If (rscuestionario.Fields.Item("piernas").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DOLOR PIERNAS/BRAZOS  <input type="checkbox" name="vesicula" value="yes" <%If (rscuestionario.Fields.Item("vesicula").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>VESICULA       <input type="checkbox" name="menospausia" value="yes" <%If (rscuestionario.Fields.Item("menospausia").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>MENOSPAUSIA
<input type="checkbox" name="tension" value="yes" <%If (rscuestionario.Fields.Item("tension").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>TENSION MUSCULAR      <input type="checkbox" name="circulacion" value="yes" <%If (rscuestionario.Fields.Item("circulacion").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>CIRCULACION    <input type="checkbox" name="parasito" value="yes" <%If (rscuestionario.Fields.Item("parasito").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PARASITO
<input type="checkbox" name="espaldas" value="yes" <%If (rscuestionario.Fields.Item("espaldas").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DOLOR DE ESPALDAS     <input type="checkbox" name="corazon" value="yes" <%If (rscuestionario.Fields.Item("corazon").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>CORAZON        <input type="checkbox" name="alta" value="yes" <%If (rscuestionario.Fields.Item("alta").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PRESION ALTA
<input type="checkbox" name="huesos" value="yes" <%If (rscuestionario.Fields.Item("huesos").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>DOLOR DE HUESOS       <input type="checkbox" name="varices" value="yes" <%If (rscuestionario.Fields.Item("varices").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>VARICES        <input type="checkbox" name="baja" value="yes" <%If (rscuestionario.Fields.Item("baja").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PRESION BAJA
<input type="checkbox" name="piel" value="yes" <%If (rscuestionario.Fields.Item("piel").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PROBLEMAS EN LA PIEL:EZEMA,PSORIOSIS,ACNE,PIEL SECA/GRASOSA
<input type="checkbox" name="cabello" value="yes" <%If (rscuestionario.Fields.Item("cabello").Value = true) Then Response.Write("CHECKED") : Response.Write("")%>>PROBLEMAS EN EL CABELLO:SE CAE MUCHO,MUY SECO,GRASOSO,CASPA,SEBORREA

De acuerdo a lo indicado anteriormente, que es lo mas importante?
<b><%=(rscuestionario.Fields.Item("importante").Value)%></b>
</pre>
              <%Session("svid")= (rscuestionario.Fields.Item("id").Value) %> 
              <input type="submit" name="Submit" value="BORRAR">
              <input type="hidden" name="MM_recordId" value="<%= rscuestionario.Fields.Item("ID").Value %>">
              <input type="hidden" name="MM_delete" value="true">
              &nbsp;&nbsp;&nbsp;<a href="cuestionarioall.asp">CANCELAR</a> 
            </form>
          </td>
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
