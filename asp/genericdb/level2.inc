<% 
If session("allow") = False Then Response.Redirect "default.asp"
If session("clearance") < 2 Then Response.Redirect "utility.asp?method=unauthorized"
%>
