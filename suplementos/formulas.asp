<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level0.inc"-->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#E7E7CF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td><%
Set rs = Server.CreateObject("ADODB.RecordSet")
nombre = Request.Form("nombre")
q= "SELECT * FROM suplementos WHERE ((suplementos.desc LIKE '" & nombre & "%') and (TYPE='P')) ORDER BY Suplementos.DESC"
rs.Open q, "DSN=7598.suplementos;"

   %> 
      <p><%if rs.EOF then%><font face="Verdana, Arial, Helvetica, sans-serif">No 
        hay ningun suplemento con ese nombre</font><%else%> 
      </p>
      <table border="0" width="100%">
        <tr> 
          <td width="20">&nbsp;</td>
          <td height="30"> 
            <p><font face="Arial, Helvetica, sans-serif"><font color="#800040" face="Verdana, Arial, Helvetica, sans-serif"><b>SUPLEMENTOS</b></font><b><font color="#800040"> 
              </font></b></font> </p>
          </td>
        </tr>
        <% while NOT rs.EOF %> 
        <tr> 
          <td width="20" align="right"> 
            <p><img src="/images/bullet.gif" width="6" height="6"><font face="Verdana, Arial, Helvetica, sans-serif"> 
              </font> 
          </td>
          <td> 
            <p><a href="formid.asp?id=<%=RS("id")%>" target="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("desc")%></font></a> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> </font> 
          </td>
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
      <p><%end if%></p>
    </td>
  </tr>
</table>
</body>
</html>
