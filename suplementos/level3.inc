<% 
If session("allow") = False Then Response.Redirect "default.asp"
If session("clearance") < 3 Then Response.Redirect "utility.asp?method=unauthorized" 'only a level three
'can authorize a new member
%>
