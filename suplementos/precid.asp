<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level1.inc"-->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td height="20"><font face="Arial, Helvetica, sans-serif"><b></b></font></td>
  </tr>
  <tr align="center"> 
    <td><%
Set rs = Server.CreateObject("ADODB.RecordSet")
q= "SELECT * FROM suplementos WHERE  id=" & Request("id")
rs.Open q, "DSN=7598.suplementos;"

DESC = rs("desc").value   %> 
      <table border="0" width="100%">
        <tr> 
          <td width="20">&nbsp;</td>
          <td> 
            <p><font face="Arial, Helvetica, sans-serif"><font face="Arial" color="#004080"><i><strong><%=desc%></strong></i></font></font></p>
            <table width="300" border="1" bgcolor="#F2F2E6">
              <tr> 
                <td width="100"><font face="Verdana, Arial, Helvetica, sans-serif"><%=rs("size")%> 
                  &nbsp;<%=rs("shape")%></font></td>
                <td width="100"><font face="Verdana, Arial, Helvetica, sans-serif"><%=rs("cont")%> 
                  </font></td>
                <td width="100"><font face="Verdana, Arial, Helvetica, sans-serif">$ 
                  <%=rs("cost")%></font></td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("indications")%></font></p>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("recetas")%></font><br>
            </p>
            <table width="100%" border="2" cellpadding="10">
              <tr> 
                <td width="150" align="center"><%if trim(rs("label")) <>"" then%><img src="../fotos/<%=rs("label")%>"><font size="2"><%end if%></font><font size="2">&nbsp; 
                  </font></td>
                <td> 
                  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%if rs("serving") <>"" then%> 
                    <%=rs("serving")%><%end if%></font></p>
                  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("formula")%></font></p>
                  <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                  <%=rs("other")%></font> 
                  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                    <%=rs("additional")%></font></p>
                </td>
              </tr>
            </table>
            <br>
            <table width="300" border="1">
              <tr> 
                <td width="150"><font face="Verdana, Arial, Helvetica, sans-serif"><%=rs("LOT")%> 
                  &nbsp;</font></td>
                <td width="100"><font face="Verdana, Arial, Helvetica, sans-serif"><%=rs("EXPDATE")%> 
                  </font></td>
              </tr>
            </table>
            <p><font size="2"> </font></p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
