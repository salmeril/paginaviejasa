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
    <td height="20"><font face="Arial, Helvetica, sans-serif"><b></b></font></td>
  </tr>
  <tr align="center"> 
    <td>
      <%
Set rs = Server.CreateObject("ADODB.RecordSet")
q= "SELECT CENTRAL.ID, CENTRAL.DESC, CENTRAL.SKU, NEWYORK.QTY as NYQTY, NEWYORK.MTD as NYMTD, NEWYORK.FIRST as NYFIRST, NEWYORK.SECON as NY2, NEWYORK.THIRD as NY3, NEWYORK.FOURT as NY4, NEWYORK.FIFTH as NY5, NEWYORK.SIXTH as NY6, MARYLAND.QTY as MAQTY, MARYLAND.MTD as MAMTD, MARYLAND.FIRST as MAFIRST, MARYLAND.SECON as MA2, MARYLAND.THIRD as MA3, MARYLAND.FOURT as MA4, MARYLAND.FIFTH as MA5, MARYLAND.SIXTH as MA6, HEMPSTEAD.QTY as HEQTY, HEMPSTEAD.MTD as HEMTD, HEMPSTEAD.FIRST as HEFIRST, HEMPSTEAD.SECON as HE2, HEMPSTEAD.THIRD as HE3, HEMPSTEAD.FOURT as HE4, HEMPSTEAD.FIFTH as HE5, HEMPSTEAD.SIXTH as HE6, DALLAS.QTY as DAQTY, DALLAS.MTD as DAMTD, DALLAS.FIRST as DAFIRST, DALLAS.SECON as DA2, DALLAS.THIRD as DA3, DALLAS.FOURT as DA4, DALLAS.FIFTH as DA5, DALLAS.SIXTH as DA6, CHICAGO.QTY as CHQTY, CHICAGO.MTD as CHMTD, CHICAGO.FIRST as CHFIRST, CHICAGO.SECON as CH2, CHICAGO.THIRD as CH3, CHICAGO.FOURT as CH4, CHICAGO.FIFTH as CH5, CHICAGO.SIXTH as CH6, WAREHOUSE.QTY as WAQTY, WAREHOUSE.MTD as WAMTD, WAREHOUSE.FIRST as WAFIRST, WAREHOUSE.SECON as WA2, WAREHOUSE.THIRD as WA3, WAREHOUSE.FOURT as WA4, WAREHOUSE.FIFTH as WA5, WAREHOUSE.SIXTH as WA6 FROM BOSTON RIGHT JOIN (WAREHOUSE RIGHT JOIN (NEWYORK RIGHT JOIN (CHICAGO RIGHT JOIN (DALLAS RIGHT JOIN (HEMPSTEAD RIGHT JOIN (MARYLAND RIGHT JOIN CENTRAL ON MARYLAND.SKU = CENTRAL.SKU) ON HEMPSTEAD.SKU = CENTRAL.SKU) ON DALLAS.SKU = CENTRAL.SKU) ON CHICAGO.SKU = CENTRAL.SKU) ON NEWYORK.SKU = CENTRAL.SKU) ON WAREHOUSE.SKU = CENTRAL.SKU) ON BOSTON.SKU = CENTRAL.SKU WHERE CENTRAL.ID=" & Request("id")
rs.Open q, "DSN=7598.inventarios;"


 %>
      <table border="0" width="100%">
        <tr> 
          <td width="20">&nbsp;</td>
          <td> 
            <p><font face="Arial, Helvetica, sans-serif"><font face="Arial" color="#004080"><i><strong><%=rs("DESC")%></strong></i></font></font></p>
            <table width="600" border="1">
              <tr bgcolor="#CCCC99"> 
                <td height="30"><font color="#000000"><b></b></font></td>
                <td align="center" height="30"><font color="#000000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Cantidad</font></b></font></td>
                <td align="center" height="30"><font color="#000000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">MTD</font></b></font></td>
                <td align="center" height="30"><font color="#000000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">1</font></b></font></td>
                <td align="center" height="30"><font color="#000000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">2</font></b></font></td>
                <td align="center" height="30"><font color="#000000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">3</font></b></font></td>
                <td align="center" height="30"><font color="#000000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">4</font></b></font></td>
                <td align="center" height="30"><font color="#000000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">5</font></b></font></td>
                <td align="center" height="30"><font color="#000000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">6</font></b></font></td>
              </tr>
              <tr> 
                <td width="200"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEW 
                  YORK</font></b></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("NYQTY")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("NYMTD")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("NYFIRST")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("NY2")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("NY3")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("NY4")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("NY5")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("NY6")%></font></td>
              </tr>
              <tr> 
                <td width="200"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">CHICAGO</font></b></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("CHQTY")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("CHMTD")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("CHFIRST")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("CH2")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("CH3")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("CH4")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("CH5")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("CH6")%></font></td>
              </tr>
              <tr> 
                <td width="200"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">MARYLAND</font></b></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MAQTY")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MAMTD")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MAFIRST")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MA2")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MA3")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MA4")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MA5")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MA6")%></font></td>
              </tr>
              <tr> 
                <td width="200"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">DALLAS</font></b></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("DAQTY")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("DAMTD")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("DAFIRST")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("DA2")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("DA3")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("DA4")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("DA5")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("DA6")%></font></td>
              </tr>
              <tr> 
                <td width="200"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">HEMPSTEAD</font></b></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("HEQTY")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("HEMTD")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("HEFIRST")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("HE2")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("HE3")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("HE4")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("HE5")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("HE6")%></font></td>
              </tr>
              <tr> 
                <td width="200"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">WAREHOUSE</font></b></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("WAQTY")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("WAMTD")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("WAFIRST")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("WA2")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("WA3")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("WA4")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("WA5")%></font></td>
                <td width="100" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("WA6")%></font></td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Updated 
              on 
              <!-- #BeginDate format:En2 -->22-Sep-2003<!-- #EndDate -->
              </font></p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
