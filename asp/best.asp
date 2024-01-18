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
q= "SELECT CENTRAL.ID, CENTRAL.DESC, NEWYORK.MTD as NYMTD, MARYLAND.MTD as MAMTD, HOUSTON.MTD as HOMTD, HEMPSTEAD.MTD as HEMTD, DALLAS.MTD as DAMTD, CHICAGO.MTD as CHMTD, BOSTON.MTD as BOMTD, [NEWYORK].[MTD]+[MARYLAND].[MTD]+[HOUSTON].[MTD]+[HEMPSTEAD].[MTD]+[DALLAS].[MTD]+[CHICAGO].[MTD]+[BOSTON].[MTD] AS Expr1 FROM BOSTON RIGHT JOIN (WAREHOUSE RIGHT JOIN (NEWYORK RIGHT JOIN (CHICAGO RIGHT JOIN (DALLAS RIGHT JOIN (HEMPSTEAD RIGHT JOIN (HOUSTON RIGHT JOIN (MARYLAND RIGHT JOIN CENTRAL ON MARYLAND.SKU = CENTRAL.SKU) ON HOUSTON.SKU = CENTRAL.SKU) ON HEMPSTEAD.SKU = CENTRAL.SKU) ON DALLAS.SKU = CENTRAL.SKU) ON CHICAGO.SKU = CENTRAL.SKU) ON NEWYORK.SKU = CENTRAL.SKU) ON WAREHOUSE.SKU = CENTRAL.SKU) ON BOSTON.SKU = CENTRAL.SKU ORDER BY [NEWYORK].[MTD]+[MARYLAND].[MTD]+[HOUSTON].[MTD]+[HEMPSTEAD].[MTD]+[DALLAS].[MTD]+[CHICAGO].[MTD]+[BOSTON].[MTD] DESC"
rs.Open q, "DSN=inventarios;"

 %> 
     
      <p><%if rs.EOF then%><font face="Verdana, Arial, Helvetica, sans-serif">No 
        hay ningun suplemento con ese nombre</font>&nbsp; <%else%> 
      </p>
      <table border="1" width="900" align="center">
        <% while NOT rs.EOF %> 
        <tr> 
          <td width="340"> 
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("desc")%></font> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> </font> 
          </td>
          <td width="70" bgcolor="#FFFFCC"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("expr1")%></font></b></div>
          </td>
          <td width="70"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("nymtd")%></font></div>
          </td>
          <td width="70"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("chmtd")%></font></div>
          </td>
          <td width="70"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("mamtd")%></font></div>
          </td>
          <td width="70"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("damtd")%></font></div>
          </td>
          <td width="70"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("bomtd")%></font></div>
          </td>
          <td width="70"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("hemtd")%></font></div>
          </td>
          <td width="70"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("homtd")%></font></div>
          </td>
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
