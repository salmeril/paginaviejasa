<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level0.inc"-->

<%
Response.Write "Welcome..."
Response.Write "<p>You have been logged <b>In</b>."
Response.Write "<br>Your clearance level is <b>" & Session("Clearance") & "</b>."
Response.Write "<p><a href=""utility.asp?method=abandon"">Log Out</a>"
%>

<html>
<body>

<p>
<br>
<br>
<br>
<br>

<center>
    <font face="Verdana, Arial, Helvetica, sans-serif"><a href="formulas.htm">FORMULAS 
    </a></font><a href="formulas.htm"></a> 
  </center>

</body>
</html>
