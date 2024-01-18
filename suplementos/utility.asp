<%@Language=VBScript%>
<%Response.Buffer=True%>

<%
method = Request.QueryString("method")

Select Case method
  Case "unauthorized"
    Unauthorized()
  Case "expired"
    Expired()
  Case "abandon"
    Abandon()
End Select


Sub Unauthorized()
  Response.Write "You do not have the security clearance to do this!"
  Response.Write "<p><a href=""default.asp"">Please Go Back</a>"
End Sub

Sub Expired()
  Response.Write "Your Account has expired!"
  Response.Write "<p>Please contact site administrator."
  Session.Abandon
End Sub

Sub Abandon()
  Session.Abandon
  Response.Redirect "default.asp"
End Sub
%>
