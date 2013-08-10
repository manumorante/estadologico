<%
if ""& session("usuario") <> "1" and ""& session("usuario") <> "2" then
	Response.Redirect("privado.asp")
end if
%>