<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/new.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Nuestra Medicina</title>
<!-- InstanceEndEditable -->
<meta name="keywords" content="" />
<meta name="description" content="" />
<link href="new/styles.css" rel="stylesheet" type="text/css" />



		<script type="text/javascript" src="new/lib/jquery-1.3.2.min.js"></script>
		<script type="text/javascript" src="new/lib/jquery.easing.1.3.js"></script>
		<script type="text/javascript" src="new/lib/jquery.coda-slider-2.0.js"></script>
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
	<a href="new/index.html" class="but" title="">Home</a><div class="but_razd"></div>
	<a href="new/enfermedades.asp" class="but" title="">Enfermedades</a><div class="but_razd"></div>
	<a href="new/articulos.asp" class="but" title="">Ultimas Novedades</a><div class="but_razd"></div>
	<a href="new/preguntas.asp" class="but" title="">Preguntas&nbsp;Frequentes</a><div class="but_razd"></div>
	<a href="new/recetascaseras.asp" class="but" title="">Recetas&nbsp;Caseras</a><div class="but_razd"></div>
	        <a href="new/drarmoza.html" class="but" title="">Cesar&nbsp;Armoza</a>




          </div>
 		<div id="col1">
		   <div id="logo"> <img src="new/images/logo.gif" />
            <h2><a href="#"><small>Complementaria- Alternativa - Natural - Acupuntura</small></a></h2>
			</div>
		</div>
		<div id="col2">
			<div id="rightlogo">
              <h2><font color="#FFFFFF">CONSULTA GRATIS</font> <br />
                <font color="#FFFFFF">CON CESAR ARMOZA</font></h2>
			  <h2><img src="new/images/email.gif" /><a href="new/email.html"> &nbsp;por&nbsp;email</a></h2>
              <br />
              <h2><img src="new/images/te.gif" />&nbsp; llame&nbsp;ahora<br/>
                </a>
                1-800-522-7099<br />
                1-718-651-6677</h2>
			</div>
		</div>


</div>

        <!-- header -->
        <!-- InstanceBeginEditable name="top" -->

						<div class="email_box">

              <table width="90%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td> <p align="left"><font color="#DBDECF" face="Verdana, Arial, Helvetica, sans-serif" size="2">
                    <%

name = Request.Form("name")
email = Request.Form("email")
telefono = Request.Form("telefono")
ciudad = Request.Form("ciudad")
pregunta = Request.Form("pregunta")

Set objMail = CreateObject("CDONTS.Newmail")
objMail.From = email
objMail.To = "cesararmoza@yahoo.com"
objMail.Subject = "Preguntele a Cesar Armoza"

sHTML =  pregunta & VbCrLf & VbCrLf
sHTML = sHTML &  name & " de " & ciudad & VbCrLf
sHTML = sHTML & telefono

objMail.Body = sHTML
objMail.Send
  %>


<%
email = Request.Form("email")
ciudad = Request.Form("ciudad")

Set objMail = CreateObject("CDONTS.Newmail")
objMail.From = email
objMail.To = "subscriptores@armonia-natural.com"
objMail.Subject = "subscribir"

objMail.Body =  ciudad
objMail.Send
  %>
                    </font></p>
                  <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Su
                    pregunta ha sido enviada a Cesar Armoza !</font><b><br>
                    </b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>
                    </font> </p>
                  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>
                    </b></font></p>
                  </td>
              </tr>
            </table>
				</div>

			  <!-- InstanceEndEditable -->

<!-- content -->



		<div id="content">



          <div class="col"> <a href="new/ejercicios.html"><img src="new/images/ejercicios.gif" border="0" /></a></div>
         	<div class="float_l" ><img src="new/images/blueline.jpg" /></div>

          <div class="col"> <font color="#CC3300"></font> <img src="new/images/hable.gif" width="150" height="22" />
            <object type="application/x-shockwave-flash" data="https://clients4.google.com/voice/embed/webCallButton" width="230" height="85">
              <param name="movie" value="https://clients4.google.com/voice/embed/webCallButton" />
              <param name="wmode" value="transparent" />
              <param name="FlashVars" value="id=032e171ff691d99aaacdcfd0f29f64a8b8d3856d&style=0" />
            </object>
            <font color="#CC3300"><br />
            </font></div>
			<div class="float_l"><img src="new/images/blueline.jpg" /></div>

          <div class="col"> <a href="new/sinnom.asp?nombre=A"><img src="new/images/indice.gif" border="0" /></a></div>
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
		   <a href="http://www.facebook.com/cesararmoza"><img src="new/images/facebook.jpg" width="98" height="32" border="0" /></a>
          </div>

          <div id="b_col4">
		  <h1>&nbsp;</h1>
		  <a href="http://www.youtube.com/cesararmoza"><img src="new/images/youtube.gif" width="87" height="36" border="0" /></a>
          </div>
			<div style="clear: both"></div>
		</div>

        <!-- / bottom -->
        <!-- footer -->
        <div id="footer_box">
		</div>
        <!-- / footer -->
        <div align="center"><a href="new/disclaimer.html">DISCLAIMER</a></div>
      </div>


</div>
</div>
</div>

</body>
<!-- InstanceEnd --></html>