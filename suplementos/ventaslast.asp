<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level2.inc"-->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td><%
Set rs = Server.CreateObject("ADODB.RecordSet")
q= "SELECT CENTRAL.ID, CENTRAL.DESC, NEWYORK.FIRST as NYFIRST, MARYLAND.FIRST as MAFIRST, HEMPSTEAD.FIRST as HEFIRST, DALLAS.FIRST as DAFIRST, CHICAGO.FIRST as CHFIRST, [NEWYORK].[FIRST]+[MARYLAND].[FIRST]+[HEMPSTEAD].[FIRST]+[DALLAS].[FIRST]+[CHICAGO].[FIRST] AS Expr1 FROM WAREHOUSE RIGHT JOIN (NEWYORK RIGHT JOIN (CHICAGO RIGHT JOIN (DALLAS RIGHT JOIN (HEMPSTEAD RIGHT JOIN (MARYLAND RIGHT JOIN CENTRAL ON MARYLAND.SKU = CENTRAL.SKU) ON HEMPSTEAD.SKU = CENTRAL.SKU) ON DALLAS.SKU = CENTRAL.SKU) ON CHICAGO.SKU = CENTRAL.SKU) ON NEWYORK.SKU = CENTRAL.SKU) ON WAREHOUSE.SKU = CENTRAL.SKU WHERE (CENTRAL.TYPE='P') ORDER BY [NEWYORK].[FIRST]+[MARYLAND].[FIRST]+[HEMPSTEAD].[FIRST]+[DALLAS].[FIRST]+[CHICAGO].[FIRST] DESC"
rs.Open q, "DSN=7598.inventarios;"

 %> 
     
      <p><%if rs.EOF then%><font face="Verdana, Arial, Helvetica, sans-serif">No 
        hay ningun suplemento con ese nombre</font>&nbsp; <%else%> 
      </p>
      <table border="1" width="786" align="left">
        <% while NOT rs.EOF %>
        <tr> 
          <td width="320"> <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("desc")%></font> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> </font> 
          </td>
          <td width="70" bgcolor="#FFFFCC"> <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("expr1")%></font></b></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("nyFIRST")%></font></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("chFIRST")%></font></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("maFIRST")%></font></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("daFIRST")%></font></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("heFIRST")%></font></div></td>
        </tr>
        <%
                rs.MoveNext
                wend
   
                %>
      </table>
      <p>&nbsp;</p>
      <p><%end if%></p>
    </td>
  </tr>
</table>
</body>
</html>
