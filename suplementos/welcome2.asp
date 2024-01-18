<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level2.inc"-->

<%
Response.Write "Welcome..."
Response.Write "<p>You have been logged <b>In</b>."
Response.Write "<br>Your clearance level is <b>" & Session("Clearance") & "</b>."
Response.Write "<p><a href=""utility.asp?method=abandon"">Log Out</a>"
%>

<html>
<body>
<p align="left"> <br>
  <br>
<table width="600" border="0" align="center">
  <tr> 
    <td width="600"> 
      <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><a href="preciosyformulas.htm"> 
        FORMULAS Y COSTOS</a></font> 
    </td>
  </tr>
  <tr>
    <td width="600"> 
      <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><a href="inventario.htm"> 
        CANTIDADES Y VENTAS DE LOS ULTIMOS 6 MESES</a></font><a href="formulas.htm"></a>
    </td>
  </tr>
  <tr> 
    <td width="600"> 
      <p align="left"> 
      <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><a href="inventario1.htm">INVENTARIO 
        DE TODAS LAS CLINICAS</a></font> </p>
    </td>
  </tr>
  <tr> 
    <td width="600">&nbsp;</td>
  </tr>
  <tr> 
    <td width="600"> 
      <p align="left"> 
      <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><b> VENTAS</b></font></p>
    </td>
  </tr>
  <tr> 
    <td width="600"><font face="Verdana, Arial, Helvetica, sans-serif"><a href="ventasmtd.htm">MTD</a></font></td>
  </tr>
  <tr> 
    <td width="600"><font face="Verdana, Arial, Helvetica, sans-serif"><a href="ventaslast.htm">LAST 
      MONTH</a></font> </td>
  </tr>
  <tr> 
    <td width="600"><font face="Verdana, Arial, Helvetica, sans-serif"><a href="ventasytd.htm">YTD</a></font></td>
  </tr>
  <tr> 
    <td width="600">&nbsp;</td>
  </tr>
  <tr> 
    <td width="600">&nbsp;</td>
  </tr>
  <tr> 
    <td width="600">&nbsp;</td>
  </tr>
</table>
<p align="left"><br>
  <br>
  <a href="formulas.htm"></a><font face="Verdana, Arial, Helvetica, sans-serif"><a href="ventaslast.htm"> 
  </a></font>
</body>
</html>
