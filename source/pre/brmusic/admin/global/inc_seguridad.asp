<%
if unerror then
	Response.Write "<b>Error</b>:<br>"& msgerror
else
	if ""&session("usuario") = "" or ""&session("cualid") = "" then
		Response.Redirect("../usuarios/validar.asp?msg=No hay una sesión o se ha perdido por inactividad.")
	elseif not getPermiso(session("cualid"),session("idioma")) then
		Response.Redirect("../usuarios/validar.asp?msg=No tiene permiso para administrar esta zona.")
	else
		' OK
	end if
end if	
%>