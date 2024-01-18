<%@LANGUAGE="VBSCRIPT"%> <%
' *** Insert Record: construct a sql insert statement and execute it
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
If (CStr(Request("MM_insert")) <> "") Then

  MM_tableName = "cuestionarios"
  MM_fields = "fecha,fecha,#,none,NULL,respondido,respondido,',none,'',nombre,nombre,',none,'',apellido,apellido,',none,'',edad,edad,none,none,NULL,direccion,direccion,',none,'',ciudad,ciudad,',none,'',estado,estado,',none,'',codigo,codigo,',none,'',lugar,lugar,',none,'',telefono,telefono,',none,'',email,email,',none,'',contacto,contacto,',none,'',horario,horario,',none,'',pregunta,pregunta,',none,'',salud,salud,',none,'',tratamiento,tratamiento,',none,'',desde,desde,',none,'',medicinas,medicinas,',none,'',vitaminas,vitaminas,',none,'',alergico,alergico,',none,'',alergias,alergias,',none,'',infeccion,infeccion,',none,'',infecciones,infecciones,',none,'',operado,operado,',none,'',operaciones,operaciones,',none,'',hiv,hiv,none,1,0,aid,aid,none,1,0,hepatitis,hepatitis,none,1,0,venereas,venereas,none,1,0,vista,vista,none,1,0,memoria,memoria,none,1,0,insomnio,insomnio,none,1,0,oido,oido,none,1,0,colesterol,colesterol,none,1,0,diabetes,diabetes,none,1,0,nariz,nariz,none,1,0,nervios,nervios,none,1,0,anemia,anemia,none,1,0,bronquitis,bronquitis,none,1,0,depresion,depresion,none,1,0,asma,asma,none,1,0,catarro,catarro,none,1,0,ansiedad,ansiedad,none,1,0,alcoholismo,alcoholismo,none,1,0,acidez,acidez,none,1,0,cansancio,cansancio,none,1,0,drogadiccion,drogadiccion,none,1,0,estomago,estomago,none,1,0,migrana,migrana,none,1,0,prostata,prostata,none,1,0,cabeza,cabeza,none,1,0,gases,gases,none,1,0,ovarios,ovarios,none,1,0,apetito,apetito,none,1,0,rinones,rinones,none,1,0,vaginal,vaginal,none,1,0,menstruales,menstruales,none,1,0,orina,orina,none,1,0,diarrea,diarrea,none,1,0,orinar,orinar,none,1,0,estrenimiento,estrenimiento,none,1,0,embarazo,embarazo,none,1,0,encias,encias,none,1,0,reumatismo,reumatismo,none,1,0,impotencia,impotencia,none,1,0,higado,higado,none,1,0,artritis,artritis,none,1,0,frigidez,frigidez,none,1,0,piernas,piernas,none,1,0,vesicula,vesicula,none,1,0,menospausia,menospausia,none,1,0,tension,tension,none,1,0,circulacion,circulacion,none,1,0,parasito,parasito,none,1,0,espaldas,espaldas,none,1,0,corazon,corazon,none,1,0,alta,alta,none,1,0,huesos,huesos,none,1,0,varices,varices,none,1,0,baja,baja,none,1,0,piel,piel,none,1,0,cabello,cabello,none,1,0,importante,importante,',none,''"
  MM_redirectPage = "cuestionarioconf.asp"

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
<%
set rscuestionario = Server.CreateObject("ADODB.Recordset")
rscuestionario.ActiveConnection = "dsn=5084.quotes;"
rscuestionario.Source = "SELECT * FROM cuestionarios"
rscuestionario.CursorType = 0
rscuestionario.CursorLocation = 2
rscuestionario.LockType = 3
rscuestionario.Open
rscuestionario_numRows = 0
%> 
<html><!-- #BeginTemplate "/Templates/basic.dwt" --><!-- DW6 -->

<head>
<!-- #BeginEditable "doctitle" --> 
<title>Untitled Document</title>
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
          <td width="15" height="29">&nbsp;</td>
          <td height="29"><b><font color="#333399">CUESTIONARIO DE SALUD</font></b></td>
        </tr>
        <tr> 
          <td width="15">&nbsp;</td>
          <td> 
            <form name="form1" method="post" action="<%=MM_editAction%>">
              <pre><input type="hidden" name="fecha" value="<%= date() %>"><input type="hidden" name="respondido" value="no">Nombre:   <input type="text" name="nombre" size="50"><br>Apellido: <input type="text" name="apellido" size="50"><br>Edad:     <input type="text" name="edad"><br>Dirección:<input type="text" name="direccion" size="50"><br>Ciudad:   <input type="text" name="ciudad">
Estado:   <input type="text" name="estado">  Código: <input type="text" name="codigo"><br>Lugar de Residencia: <select name="lugar"><option value="USA" selected>USA</option><option value="AMERICA">CENTRO Y SUD AMERICA</option><option value="EUROPA">EUROPA</option><option value="OCEANIA">OCEANIA</option><option value="ASIA">ASIA</option><option value="AFRICA">AFRICA</option></select><br>Teléfono: <input type="text" name="telefono"><br><b>Email:</b>    <input type="text" name="email" size="50">

Como desea ser contactado? <select name="contacto" size="1"><option selected>por email</option><option>por teléfono</option></select><br>Horario conveniente para llamarlo <input type="text" name="horario">

<b>Problema principal y pregunta?</b> <br><textarea name="pregunta" cols="50" rows="3"></textarea>

Que le dijo su doctor acerca de su salud? <input type="text" name="salud" size="50">
Esta recibiendo tratamiento medico?  <input type="radio" name="tratamiento" value="si">si  <input type="radio" name="tratamiento" value="no" checked>no
Desde cuando? <input type="text" name="desde" size="50">
Que medicinas esta tomando? <input type="text" name="medicinas" size="50">
Que vitaminas esta tomando? <input type="text" name="vitaminas" size="50">
Es usted alérgico a algun medicamento?  <input type="radio" name="alergico" value="si">si  <input type="radio" name="alergico" value="no" checked>no
Cual? <input type="text" name="alergias" size="50">
Alguna infección?  <input type="radio" name="infeccion" value="si">si  <input type="radio" name="infeccion" value="no" checked>no
Cual? <input type="text" name="infecciones" size="50">
Ha sido operado?  <input type="radio" name="operado" value="si">si  <input type="radio" name="operado" value="no" checked>no
De que? <input type="text" name="operaciones" size="50">

Tiene o ha tenido:

<input type="checkbox" name="hiv" value="yes">HIV   <input type="checkbox" name="aid" value="yes">AID   <input type="checkbox" name="hepatitis" value="yes">HEPATITIS    <input type="checkbox" name="venereas" value="yes">ENFERMEDADES VENEREAS
<input type="checkbox" name="vista" value="yes">PROBLEMAS DE VISTA    <input type="checkbox" name="memoria" value="yes">MEMORIA        <input type="checkbox" name="insomnio" value="yes">INSOMNIO
<input type="checkbox" name="oido" value="yes">PROBLEMAS DE OIDO     <input type="checkbox" name="colesterol" value="yes">COLESTEROL     <input type="checkbox" name="diabetes" value="yes">DIABETES
<input type="checkbox" name="nariz" value="yes">ALERGIAS NARIZ Y PIEL <input type="checkbox" name="nervios" value="yes">NERVIOS        <input type="checkbox" name="anemia" value="yes">ANEMIA
<input type="checkbox" name="bronquitis" value="yes">BRONQUITIS            <input type="checkbox" name="depresion" value="yes">DEPRESION      <input type="checkbox" name="asma" value="yes">ASMA
<input type="checkbox" name="catarro" value="yes">MUCHO CATARRO         <input type="checkbox" name="ansiedad" value="yes">ANSIEDAD       <input type="checkbox" name="alcoholismo" value="yes">ALCOHOLISMO
<input type="checkbox" name="acidez" value="yes">ACIDEZ ESTOMACAL      <input type="checkbox" name="cansancio" value="yes">CANSANCIO      <input type="checkbox" name="drogadiccion" value="yes">DROGADICCION
<input type="checkbox" name="estomago" value="yes">DOLOR DE ESTOMAGO     <input type="checkbox" name="migrana" value="yes">MIGRAÑA        <input type="checkbox" name="prostata" value="yes">PROSTATA
<input type="checkbox" name="cabeza" value="yes">DOLORES DE CABEZA     <input type="checkbox" name="gases" value="yes">GASES          <input type="checkbox" name="ovarios" value="yes">OVARIOS
<input type="checkbox" name="apetito" value="yes">APETITO MUCHO POCO    <input type="checkbox" name="rinones" value="yes">RIÑONES        <input type="checkbox" name="vaginal" value="yes">INFECCION VAGINAL
<input type="checkbox" name="menstruales" value="yes">PROBLEMAS MENSTRUALES <input type="checkbox" name="orina" value="yes">ORINA MUCHO    <input type="checkbox" name="diarrea" value="yes">DIARREA
<input type="checkbox" name="orinar" value="yes">MOLESTIAS AL ORINAR   <input type="checkbox" name="estrenimiento" value="yes">ESTREÑIMIENTO  <input type="checkbox" name="embarazo" value="yes">EMBARAZO
<input type="checkbox" name="encias" value="yes">ENCIAS IMFLAMADAS     <input type="checkbox" name="reumatismo" value="yes">REUMATISMO     <input type="checkbox" name="impotencia" value="yes">IMPOTENCIA
<input type="checkbox" name="higado" value="yes">HIGADO                <input type="checkbox" name="artritis" value="yes">ARTRITIS       <input type="checkbox" name="frigidez" value="yes">FRIGIDEZ
<input type="checkbox" name="piernas" value="yes">DOLOR PIERNAS/BRAZOS  <input type="checkbox" name="vesicula" value="yes">VESICULA       <input type="checkbox" name="menospausia" value="yes">MENOSPAUSIA
<input type="checkbox" name="tension" value="yes">TENSION MUSCULAR      <input type="checkbox" name="circulacion" value="yes">CIRCULACION    <input type="checkbox" name="parasito" value="yes">PARASITO
<input type="checkbox" name="espaldas" value="yes">DOLOR DE ESPALDAS     <input type="checkbox" name="corazon" value="yes">CORAZON        <input type="checkbox" name="alta" value="yes">PRESION ALTA
<input type="checkbox" name="huesos" value="yes">DOLOR DE HUESOS       <input type="checkbox" name="varices" value="yes">VARICES        <input type="checkbox" name="baja" value="yes">PRESION BAJA
<input type="checkbox" name="piel" value="yes">PROBLEMAS EN LA PIEL:EZEMA,PSORIOSIS,ACNE,PIEL SECA/GRASOSA
<input type="checkbox" name="cabello" value="yes">PROBLEMAS EN EL CABELLO:SE CAE MUCHO,MUY SECO,GRASOSO,CASPA,SEBORREA

De acuerdo a lo indicado anteriormente, que es lo mas importante?
<input type="text" name="importante" size="50">

<input type="submit" name="submit" value="Mandar" onClick="MM_validateForm('edad','','NisNum','email','','RisEmail','pregunta','','R');return document.MM_returnValue">
</pre>
              <input type="hidden" name="MM_insert" value="true">
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
