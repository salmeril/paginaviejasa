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
q= "SELECT CENTRAL.ID, CENTRAL.DESC, [NEWYORK].[MTD]+[NEWYORK].[FIRST]+[NEWYORK].[SECON]+[NEWYORK].[THIRD]+[NEWYORK].[FOURT]+[NEWYORK].[FIFTH]+[NEWYORK].[SIXTH]+[NEWYORK].[SEVEN]+[NEWYORK].[EIGHTH]+[NEWYORK].[NINTH]+[NEWYORK].[TENTH]+[NEWYORK].[ELEVEN]+[NEWYORK].[TWELVE] AS NYYTD, [MARYLAND].[MTD]+[MARYLAND].[FIRST]+[MARYLAND].[SECON]+[MARYLAND].[THIRD]+[MARYLAND].[FOURT]+[MARYLAND].[FIFTH]+[MARYLAND].[SIXTH]+[MARYLAND].[SEVEN]+[MARYLAND].[EIGHTH]+[MARYLAND].[NINTH]+[MARYLAND].[TENTH]+[MARYLAND].[ELEVEN]+[MARYLAND].[TWELVE] AS MAYTD, [HEMPSTEAD].[MTD]+[HEMPSTEAD].[FIRST]+[HEMPSTEAD].[SECON]+[HEMPSTEAD].[THIRD]+[HEMPSTEAD].[FOURT]+[HEMPSTEAD].[FIFTH]+[HEMPSTEAD].[SIXTH]+[HEMPSTEAD].[SEVEN]+[HEMPSTEAD].[EIGHTH]+[HEMPSTEAD].[NINTH]+[HEMPSTEAD].[TENTH]+[HEMPSTEAD].[ELEVEN]+[HEMPSTEAD].[TWELVE] AS HEYTD, [DALLAS].[MTD]+[DALLAS].[FIRST]+[DALLAS].[SECON]+[DALLAS].[THIRD]+[DALLAS].[FOURT]+[DALLAS].[FIFTH]+[DALLAS].[SIXTH]+[DALLAS].[SEVEN]+[DALLAS].[EIGHTH]+[DALLAS].[NINTH]+[DALLAS].[TENTH]+[DALLAS].[ELEVEN]+[DALLAS].[TWELVE] AS DAYTD, [CHICAGO].[MTD]+[CHICAGO].[FIRST]+[CHICAGO].[SECON]+[CHICAGO].[THIRD]+[CHICAGO].[FOURT]+[CHICAGO].[FIFTH]+[CHICAGO].[SIXTH]+[CHICAGO].[SEVEN]+[CHICAGO].[EIGHTH]+[CHICAGO].[NINTH]+[CHICAGO].[TENTH]+[CHICAGO].[ELEVEN]+[CHICAGO].[TWELVE] AS CHYTD, [BOSTON].[MTD]+[BOSTON].[FIRST]+[BOSTON].[SECON]+[BOSTON].[THIRD]+[BOSTON].[FOURT]+[BOSTON].[FIFTH]+[BOSTON].[SIXTH]+[BOSTON].[SEVEN]+[BOSTON].[EIGHTH]+[BOSTON].[NINTH]+[BOSTON].[TENTH]+[BOSTON].[ELEVEN]+[BOSTON].[TWELVE] AS BOYTD, NYYTD+MAYTD+CHYTD+HEYTD+DAYTD+BOYTD AS EXPR1 FROM BOSTON RIGHT JOIN (WAREHOUSE RIGHT JOIN (NEWYORK RIGHT JOIN (CHICAGO RIGHT JOIN (DALLAS RIGHT JOIN (HEMPSTEAD RIGHT JOIN (MARYLAND RIGHT JOIN CENTRAL ON MARYLAND.SKU = CENTRAL.SKU) ON HEMPSTEAD.SKU = CENTRAL.SKU) ON DALLAS.SKU = CENTRAL.SKU) ON CHICAGO.SKU = CENTRAL.SKU) ON NEWYORK.SKU = CENTRAL.SKU) ON WAREHOUSE.SKU = CENTRAL.SKU) ON BOSTON.SKU = CENTRAL.SKU WHERE (CENTRAL.TYPE='P') ORDER BY [NEWYORK].[MTD]+[NEWYORK].[FIRST]+[NEWYORK].[SECON]+[NEWYORK].[THIRD]+[NEWYORK].[FOURT]+[NEWYORK].[FIFTH]+[NEWYORK].[SIXTH]+[NEWYORK].[SEVEN]+[NEWYORK].[EIGHTH]+[NEWYORK].[NINTH]+[NEWYORK].[TENTH]+[NEWYORK].[ELEVEN]+[NEWYORK].[TWELVE] +[MARYLAND].[MTD]+[MARYLAND].[FIRST]+[MARYLAND].[SECON]+[MARYLAND].[THIRD]+[MARYLAND].[FOURT]+[MARYLAND].[FIFTH]+[MARYLAND].[SIXTH]+[MARYLAND].[SEVEN]+[MARYLAND].[EIGHTH]+[MARYLAND].[NINTH]+[MARYLAND].[TENTH]+[MARYLAND].[ELEVEN]+[MARYLAND].[TWELVE] + [HEMPSTEAD].[MTD]+[HEMPSTEAD].[FIRST]+[HEMPSTEAD].[SECON]+[HEMPSTEAD].[THIRD]+[HEMPSTEAD].[FOURT]+[HEMPSTEAD].[FIFTH]+[HEMPSTEAD].[SIXTH]+[HEMPSTEAD].[SEVEN]+[HEMPSTEAD].[EIGHTH]+[HEMPSTEAD].[NINTH]+[HEMPSTEAD].[TENTH]+[HEMPSTEAD].[ELEVEN]+[HEMPSTEAD].[TWELVE] + [DALLAS].[MTD]+[DALLAS].[FIRST]+[DALLAS].[SECON]+[DALLAS].[THIRD]+[DALLAS].[FOURT]+[DALLAS].[FIFTH]+[DALLAS].[SIXTH]+[DALLAS].[SEVEN]+[DALLAS].[EIGHTH]+[DALLAS].[NINTH]+[DALLAS].[TENTH]+[DALLAS].[ELEVEN]+[DALLAS].[TWELVE] + [CHICAGO].[MTD]+[CHICAGO].[FIRST]+[CHICAGO].[SECON]+[CHICAGO].[THIRD]+[CHICAGO].[FOURT]+[CHICAGO].[FIFTH]+[CHICAGO].[SIXTH]+[CHICAGO].[SEVEN]+[CHICAGO].[EIGHTH]+[CHICAGO].[NINTH]+[CHICAGO].[TENTH]+[CHICAGO].[ELEVEN]+[CHICAGO].[TWELVE]   DESC "
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
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("nyYTD")%></font></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("chYTD")%></font></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("maYTD")%></font></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("daYTD")%></font></div></td>
          <td width="70"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("heYTD")%></font></div></td>
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
