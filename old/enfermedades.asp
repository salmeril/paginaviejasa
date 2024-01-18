<?xml version="1.0" encoding="utf-8"?><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/new.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
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
	<a href="enfermedades.html" class="but" title="">Enfermedades</a><div class="but_razd"></div>
	<a href="topicos.html" class="but" title="">Topicos&nbsp;de&nbsp;Interes</a><div class="but_razd"></div>
	<a href="preguntas.html" class="but" title="">Preguntas&nbsp;Frequentes</a><div class="but_razd"></div>
	<a href="recetas.html" class="but" title="">Recetas&nbsp;Caseras</a><div class="but_razd"></div>
	        <a href="cesararmoza.html" class="but" title="">Cesar&nbsp;Armoza</a>
			<img src="images/english.jpg" width="37" height="25" /> 
          </div>
 		<div id="col1">
		   <div id="logo"> <img src="images/logo.gif" />
            <h2><a href="#"><small>Complementaria- Alternativa - Natural - Acupuntura</small></a></h2>
			</div>
		</div>
		<div id="col2">
			<div id="right">
              <h2><font color="#FFFFFF">CONSULTA GRATIS</font><br />
                <br />
              </h2>
			  <h2><img src="images/email.gif" /><a href="cesararmoza.html"> &nbsp;por&nbsp;email</a></h2>
              <br />
              <h2><img src="images/te.gif" />&nbsp; llame&nbsp;ahora<br/>
                </a>
                1-800-522-7099<br />
                1-718-651-6677</h2>
			</div>
		</div>
	
	
</div>

        <!-- header -->
        <div class="top"> <!-- InstanceBeginEditable name="top" -->
				<div id="content_blog">
					<div id="right">
					</div>
					<div id="left">
						<div class="left_box">
								
            <div class="line"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr align="center"> 
                  <td height="20"><font face="Arial, Helvetica, sans-serif"><b></b></font></td>
                </tr>
                <tr align="center"> 
                  <td>
                    <%
Set rs = Server.CreateObject("ADODB.RecordSet")
q= "SELECT * FROM enfermedades ORDER by nombre"
rs.Open q, "DSN=7598.medicina;"

   %>
                    <p>
                      <%if rs.EOF then%>
                      <font face="Verdana, Arial, Helvetica, sans-serif">No hay 
                      ninguna enfermedad</font>&nbsp; 
                      <%else%>
                    </p>
                    <table border="0" width="100%">
                      <tr> 
                        <td width="20">&nbsp;</td>
                        <td height="30"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td><font face="Arial, Helvetica, sans-serif"><b><font color="#800040">ENFERMEDADES</font></b></font></td>
                              <td align="right"><font size="2"><a href="../bibliografia.htm">BIBLIOGRAFIA</a></font></td>
                            </tr>
                          </table></td>
                      </tr>
                      <% while NOT rs.EOF %>
                      <tr> 
                        <td width="20" align="right"> <p><img src="/images/bullet.gif" width="6" height="6" /><font face="Verdana, Arial, Helvetica, sans-serif"> 
                            </font> </p></td>
                        <td> <p><a href="enfid.asp?id=<%=RS("id")%>"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
                            <%=RS("nombre")%>
                            </font></a> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                            </font> </p></td>
                      </tr>
                      <%
                rs.MoveNext
                wend
   
                %>
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
        </div>
<!-- content -->



		<div id="content">
					
						
          	<div class="col"> <img src="images/ejercicios.gif" /></div>
         	<div class="float_l" ><img src="images/blueline.jpg" /></div>
         	<div class="col"><img src="images/cesararmoza.gif" /></div>
			<div class="float_l"></div>
			<div class="col">
			<object type="application/x-shockwave-flash" data="https://clients4.google.com/voice/embed/webCallButton" width="230" height="85"><param name="movie" value="https://clients4.google.com/voice/embed/webCallButton" /><param name="wmode" value="transparent" /><param name="FlashVars" value="id=b3ba45bd4197f8d4e290aa9c68331bb2cc1fec3f&style=0" /></object>
			</div>
			<div class="float_l"><img src="images/blueline.jpg" /></div>
          	<div class="col"> <img src="images/indice.gif" /></div>
			<div style="clear: both"></div>
				
		</div>

<!-- / content --> 
		<div style="height:15px"></div>
<!-- bottom -->
		<div id="bottom">
			<div id="b_col2">
			<h1>&nbsp;</h1>
			
            <p>&nbsp;</p> <img src="images/skype.gif" width="120" height="53" />
       		</div>
        	<div id="b_col3">
            <h2>&nbsp;</h2>
			
            <p>&nbsp;</p>		<img src="images/facebook.jpg" width="130" height="45" />


        	</div>
			<div id="b_col4">
			<h1>&nbsp;</h1>
             
            <p>&nbsp; </p>
             <img src="images/youtube.gif" width="122" height="49" />
			</div>
			<div style="clear: both"></div>
		</div>
	
        <!-- / bottom -->
        <!-- footer -->
        <div id="footer_box">
		</div>
        <!-- / footer -->
        <div align="center">DISCLAIMER</div>
      </div>


</div>
</div>
</div>

</body>
<!-- InstanceEnd --></html>
