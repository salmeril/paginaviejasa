<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level2.inc"-->

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
q= "SELECT CENTRAL.ID, CENTRAL.DESC, CENTRAL.SKU, CENTRAL.TYPE,NEWYORK.QTY, NEWYORK.FIRST, MARYLAND.QTY, MARYLAND.FIRST,  HEMPSTEAD.QTY, HEMPSTEAD.FIRST, DALLAS.QTY, DALLAS.FIRST, CHICAGO.QTY, CHICAGO.FIRST, WAREHOUSE.QTY, WAREHOUSE.FIRST FROM WAREHOUSE RIGHT JOIN (NEWYORK RIGHT JOIN (CHICAGO RIGHT JOIN (DALLAS RIGHT JOIN (HEMPSTEAD RIGHT JOIN (MARYLAND RIGHT JOIN CENTRAL ON MARYLAND.SKU = CENTRAL.SKU)  ON HEMPSTEAD.SKU = CENTRAL.SKU) ON DALLAS.SKU = CENTRAL.SKU) ON CHICAGO.SKU = CENTRAL.SKU) ON NEWYORK.SKU = CENTRAL.SKU) ON WAREHOUSE.SKU = CENTRAL.SKU  WHERE (central.desc LIKE '" & nombre & "%') ORDER BY CENTRAL.DESC"
rs.Open q, "DSN=7598.inventarios;"


 %> 
     
      <p><%if rs.EOF then%><font face="Verdana, Arial, Helvetica, sans-serif">No 
        hay ningun suplemento con ese nombre</font>&nbsp; <%else%> 
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
            <p><a href="inventarioid.asp?id=<%=RS("ID")%>" target="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("desc")%></font></a> 
              <%if RS("type")<>"P" then%>(<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("type")%></font>) 
              <%end if%> 
          </td>
        </tr>
        <%
                rs.MoveNext
                wend
   
                %> 
        <tr> 
          <td width="20" align="right" height="22">&nbsp;</td>
          <td height="22">&nbsp;</td>
        </tr>
      </table>
      <p><%end if%></p>
    </td>
  </tr>
</table>
</body>
</html>
