<HR>
<FONT SIZE=-1>
<STRONG>This page was created with GenericDB.</STRONG>
<P>
<%
  curFile = server.mappath(Request.ServerVariables("PATH_INFO"))
  Set FileObj = Server.CreateObject ("LastMod.FileObj")
  FileObj.TimeFormat = "hh:mm:ss am/pm"
  FileObj.DateFormat = "ddd, mmm d yyyy"
%>
Last modified <%=FileObj.LastModified(curFile) %> <BR>
The time is now <%=Time %> EST <BR>
