<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/new.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Nuestra Medicina</title>
<!-- InstanceEndEditable -->
<meta name="keywords" content="" />
<meta name="description" content="" />
<link href="styles.css" rel="stylesheet" type="text/css" />



		<script type="text/javascript" src="lib/jquery-1.3.2.min.js"></script>
		<script type="text/javascript" src="lib/jquery.easing.1.3.js"></script>
		<script type="text/javascript" src="lib/jquery.coda-slider-2.0.js"></script>
<!-- Initialize each slider on the page. Each slider must have a unique id -->
	<script type="text/javascript">
	$().ready(function() {
	$('#coda-slider-2').codaSlider({
		autoSlide: true,
		autoSlideInterval: 6000,
		autoSlideStopWhenClicked: true

	});
 });
</script>
<!-- End JavaScript -->
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
</head>
<body>

<div id="foot_bg">
<div id="menu">
<div id="for_home">



<div id="main">
<!-- header -->
<div id="header">
	<div id="buttons">
	<a href="index.html" class="but" title="">Home</a><div class="but_razd"></div>
	<a href="enfermedades.asp" class="but" title="">Enfermedades</a><div class="but_razd"></div>
	<a href="articulos.asp" class="but" title="">Ultimas Novedades</a><div class="but_razd"></div>
	<a href="preguntas.asp" class="but" title="">Preguntas&nbsp;Frequentes</a><div class="but_razd"></div>
	<a href="recetascaseras.asp" class="but" title="">Recetas&nbsp;Caseras</a><div class="but_razd"></div>
	        <a href="drarmoza.html" class="but" title="">Cesar&nbsp;Armoza</a>




          </div>
 		<div id="col1">
		   <div id="logo"> <img src="images/logo.gif" />
            <h2><a href="#"><small>Complementaria- Alternativa - Natural - Acupuntura</small></a></h2>
			</div>
		</div>
		<div id="col2">
			<div id="rightlogo">
              <h2><font color="#FFFFFF">CONSULTA GRATIS</font> <br />
                <font color="#FFFFFF">CON CESAR ARMOZA</font></h2>
			  <h2><img src="images/email.gif" /><a href="email.html"> &nbsp;por&nbsp;email</a></h2>
              <br />
              <h2><img src="images/te.gif" />&nbsp; llame&nbsp;ahora<br/>
                </a>
                1-800-522-7099<br />
                1-718-651-6677</h2>
			</div>
		</div>


</div>

        <!-- header -->
        <!-- InstanceBeginEditable name="top" -->
				<div id="content_blog">
					<div id="right">
					</div>
					<div id="left">
						<div class="left_box">

              <div class="line">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr align="center">
                    <td height="20"> <table border="0" width="100%">
                        <tr>
                          <td width="20">&nbsp;</td>
                          <td height="30"> <p align="left"><img src="images/indice.gif" width="128" height="74" />
                            </p></td>
                        </tr>
                        <tr>
                          <td width="20" align="right">&nbsp;</td>
                          <td height="5"><div align="left"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><a href="/sinnom.asp?nombre=a">A</a>
                              <a href="/sinnom.asp?nombre=B">B</a> <a href="/sinnom.asp?nombre=C">C</a>
                              <a href="/sinnom.asp?nombre=D">D</a> <a href="/sinnom.asp?nombre=E">E</a>
                              <a href="/sinnom.asp?nombre=F">F</a> <a href="/sinnom.asp?nombre=G">G</a>
                              <a href="/sinnom.asp?nombre=H">H</a> <a href="/sinnom.asp?nombre=I">I</a>
                              <a href="/sinnom.asp?nombre=J">J</a> <a href="/sinnom.asp?nombre=K">K</a>
                              <a href="/sinnom.asp?nombre=L">L</a> <a href="/sinnom.asp?nombre=M">M</a>
                              <a href="/sinnom.asp?nombre=N">N</a> <a href="/sinnom.asp?nombre=O">O</a>
                              <a href="/sinnom.asp?nombre=P">P</a> <a href="/sinnom.asp?nombre=Q">Q</a>
                              <a href="/sinnom.asp?nombre=R">R</a> <a href="/sinnom.asp?nombre=S">S</a>
                              <a href="/sinnom.asp?nombre=T">T</a> <a href="/sinnom.asp?nombre=U">U</a>
                              <a href="/sinnom.asp?nombre=V">V</a> <a href="/sinnom.asp?nombre=W">W</a>
                              <a href="/sinnom.asp?nombre=X">X</a> <a href="/sinnom.asp?nombre=Y">Y</a>
                              <a href="/sinnom.asp?nombre=Z">Z</a></font></div></td>
                        </tr>
                        <tr>
                          <td width="20" align="right">&nbsp;</td>
                          <td height="20">&nbsp;</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr align="center">
                    <td>
                      <%
Set rs = Server.CreateObject("ADODB.RecordSet")
nombre=Request("nombre")
q= "SELECT DISTINCT nombre FROM sintomas WHERE (nombre LIKE '" & nombre & "%') ORDER BY nombre"
rs.Open q, "DSN=7598.medicina;"
%>
                      <p>
                        <%if rs.EOF then%>
                        <font face="Verdana, Arial, Helvetica, sans-serif">No
                        hay ningun sintoma</font>
                        <%else%>
                      </p>
                      <table border="0" width="100%">
                        <%while NOT rs.EOF%>
                        <tr>
                          <td width="20" align="right"> <p></p></td>
                          <td> <p align="left"><a href="sintomas.asp?nombre=<%=Server.URLEncode(RS("nombre"))%>"><%=RS("nombre")%></a> </p></td>
                        </tr>
                        <%
                rs.MoveNext
                wend

                %>
                        <tr>
                          <td width="20" align="right">&nbsp;</td>
                          <td height="5">&nbsp;</td>
                        </tr>
                      </table>
                      <p>
                        <%end if%>
                      </p></td>
                  </tr>
                </table>
                <h1>&nbsp;</h1>

							</div>
							</div>
							<div style=" height:10px"></div>


							</div>
						<div style="clear: both"></div>
				</div>

			  <!-- InstanceEndEditable -->

<!-- content -->



		<div id="content">



          <div class="col"> <a href="ejercicios.html"><img src="images/ejercicios.gif" border="0" /></a></div>
         	<div class="float_l" ><img src="images/blueline.jpg" /></div>

          <div class="col"> <font color="#CC3300"></font> <img src="images/hable.gif" width="150" height="22" />
            <object type="application/x-shockwave-flash" data="https://clients4.google.com/voice/embed/webCallButton" width="230" height="85">
              <param name="movie" value="https://clients4.google.com/voice/embed/webCallButton" />
              <param name="wmode" value="transparent" />
              <param name="FlashVars" value="id=032e171ff691d99aaacdcfd0f29f64a8b8d3856d&style=0" />
            </object>
            <font color="#CC3300"><br />
            </font></div>
			<div class="float_l"><img src="images/blueline.jpg" /></div>

          <div class="col"> <a href="sinnom.asp?nombre=A"><img src="images/indice.gif" border="0" /></a></div>
			<div style="clear: both"></div>

		</div>

<!-- / content -->
		<div style="height:15px"></div>
<!-- bottom -->
		<div id="bottom">


          <div id="b_col1">
		  <h1>&nbsp;</h1>
            <div id="google_translate_element"></div>
			<script>
function googleTranslateElementInit() {
  new google.translate.TranslateElement({
    pageLanguage: 'es',
    includedLanguages: 'en,fr,de,it,pt',
    autoDisplay: false,
    layout: google.translate.TranslateElement.InlineLayout.SIMPLE
  }, 'google_translate_element');
}
</script>
<script src="//translate.google.com/translate_a/element.js?cb=googleTranslateElementInit"></script>
       		</div>

          <div id="b_col2">
		  <h1>&nbsp;</h1>
<!--
Skype 'My status' button
http://www.skype.com/go/skypebuttons
-->
<script type="text/javascript" src="http://download.skype.com/share/skypebuttons/js/skypeCheck.js"></script>
<a href="skype:cesar.armoza ?call"><img src="http://mystatus.skype.com/smallclassic/cesar%2Earmoza " style="border: none;" width="114" height="20" alt="My status" /></a>


        	</div>


          <div id="b_col3">
		  <h1>&nbsp;</h1>
		   <a href="http://www.facebook.com/cesararmoza"><img src="images/facebook.jpg" width="98" height="32" border="0" /></a>
          </div>

          <div id="b_col4">
		  <h1>&nbsp;</h1>
		  <a href="http://www.youtube.com/cesararmoza"><img src="images/youtube.gif" width="87" height="36" border="0" /></a>
          </div>
			<div style="clear: both"></div>
		</div>

        <!-- / bottom -->
        <!-- footer -->
        <div id="footer_box">
		</div>
        <!-- / footer -->
        <div align="center"><a href="disclaimer.html">DISCLAIMER</a></div>
      </div>


</div>
</div>
</div>

</body>
<!-- InstanceEnd --></html>
