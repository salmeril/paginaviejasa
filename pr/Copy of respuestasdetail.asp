<%@LANGUAGE="VBSCRIPT"%> 
<%

Dim rsrecomendados__MMColParam
rsrecomendados__MMColParam = "0"
if(Request.QueryString("ID") <> "") then rsrecomendados__MMColParam = Request.QueryString("ID")

%> <%
set rsrecomendados = Server.CreateObject("ADODB.Recordset")
rsrecomendados.ActiveConnection = "dsn=5084.quotes;"
rsrecomendados.Source = "SELECT * FROM recomendados WHERE ID = " + Replace(rsrecomendados__MMColParam, "'", "''") + ""
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
<html><!-- #BeginTemplate "/Templates/basic.dwt" --><!-- DW6 -->

<head>
<!-- #BeginEditable "doctitle" --> 
<title>Respuesta al cuestionario</title>
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
          <td width="20">&nbsp;</td>
          <td height="5">&nbsp; </td>
        </tr>
        <tr> 
          <td width="20">&nbsp;</td>
          <td height="30"> 
            <p> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Estimada/o:<b> 
              </b><%=(rsrecomendados.Fields.Item("nombre").Value)%> 
              </font> 
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rsrecomendados.Fields.Item("contestacion").Value)%></font> 
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Venga 
              por la Clinica, llame para cita al <b>1-800-522-7099</b></font> 
              <% if (rsrecomendados.Fields.Item("total").Value)<>0 then %> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2">,o ordene 
              los siguientes productos en caso que no puede venir. </font> 
            <p> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Suplementos 
              nutricionales naturales recomendados por el Dr. Armoza <br>
              <%=(rsrecomendados.Fields.Item("productos").Value)%> 
              </font> 
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Total:<b> 
              </b>$ <%=(rsrecomendados.Fields.Item("total").Value)%></font> 
              <font face="Verdana, Arial, Helvetica, sans-serif"> <font size="2">+ 
              $10 (g</font></font><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif">astos 
              de envio)</font> </font> 
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Si 
              desea ordenarlos haga <a HREF="ordenar.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & rsrecomendados.Fields.Item("ID").Value %>">click 
              aqui</a></font><a href="../ordenar"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              </font></b> </a><% end if%> 
            </td>
        </tr>
        <tr> 
          <td width="20" align="right"> 
            <p><font face="Verdana, Arial, Helvetica, sans-serif"> </font> 
          </td>
          <td> 
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> </font> 
          </td>
        </tr>
        <tr> 
          <td width="20" align="right">&nbsp;</td>
          <td height="5">&nbsp;</td>
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
