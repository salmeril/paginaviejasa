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
						<div class="right_box">
							<div class="line">
							<%
Set rs = Server.CreateObject("ADODB.RecordSet")
q= "SELECT * FROM enfermedades WHERE  id=" & Request("id")
rs.Open q, "DSN=7598.medicina;"

  nombre = rs("nombre").value   %>


                <table width="200" border="0" cellspacing="5">
                  <tr>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"><a href="enfid.asp?id=<%=RS("id")%>">CONCEPTO</a></td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo1")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft1.asp?id=<%=RS("id")%>"><%=RS("titulo1")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo2")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft2.asp?id=<%=RS("id")%>"><%=RS("titulo2")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo3")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft3.asp?id=<%=RS("id")%>"><%=RS("titulo3")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo4")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft4.asp?id=<%=RS("id")%>"><%=RS("titulo4")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo5")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft5.asp?id=<%=RS("id")%>"><%=RS("titulo5")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo6")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft6.asp?id=<%=RS("id")%>"><%=RS("titulo6")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo7")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft7.asp?id=<%=RS("id")%>"><%=RS("titulo7")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo8")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft8.asp?id=<%=RS("id")%>"><%=RS("titulo8")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo9")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft9.asp?id=<%=RS("id")%>"><%=RS("titulo9")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo10")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft10.asp?id=<%=RS("id")%>"><%=RS("titulo10")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo11")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft11.asp?id=<%=RS("id")%>"><%=RS("titulo11")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo12")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft12.asp?id=<%=RS("id")%>"><%=RS("titulo12")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo13")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft13.asp?id=<%=RS("id")%>"><%=RS("titulo13")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo14")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft14.asp?id=<%=RS("id")%>"><%=RS("titulo14")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo15")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft15.asp?id=<%=RS("id")%>"><%=RS("titulo15")%></a> <%end if%> </td>
                  </tr>
                  <tr>
                    <%if ltrim(rs("titulo16")) <>"" then%>
                    <td align="center" width="1" bgcolor="#333366">&nbsp;</td>
                    <td width="189"> <a href="enft16.asp?id=<%=RS("id")%>"><%=RS("titulo16")%></a> <%end if%> </td>
                  </tr>
                </table>


									<div style=" height:5px"></div>
							</div>
							</div>
							<div style=" height:10px"></div>





					</div>
						<div id="left1">
							<div class="left_box1">

              <div class="line">

                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr align="center">
                    <td> <table border="0" width="100%">
                        <tr>
                          <td width="20">&nbsp;</td>
                          <td><div align="left"><b><font color="#800040">ENFERMEDADES:</font><font color="#333333">
                              </font></b><font color="#000080"><strong><font color="#000080"><strong><font color="#004080"><i><%=nombre%></i></font></strong></font><font color="#004080"></font></strong></font>
                            </div></td>
                        </tr>
                        <tr>
                          <td width="20">&nbsp;</td>
                          <td><div align="left"><b><%=RS("titulo15")%></b> </div>
                            <p align="left"><%=RS("contenido15")%> </td>
                        </tr>
                        <tr>
                          <td width="20">&nbsp;</td>
                          <td height="5">&nbsp;</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>

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
