<% 
If session("allow") = False Then Response.Redirect "default.asp"
If session("clearance") < 0 Then Response.Redirect "utility.asp?method=unauthorized"
%>
