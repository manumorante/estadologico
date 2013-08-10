<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<html>
<head>
<title>Administración</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body class="bodyAdmin">
<%if unerror then
	Response.Write "<b>Error</b><br>"&msgerror
else

	' PARA EMPEZAR A FUNCIONAR DEBEMOS SABER LO QUE DESEAMOS ADMINISTRAR
	Dim cualid ' Cualiadad, o zona a administrar. (EJ: Noticias, Archivos, ...)
	cualid = ""&request.QueryString("cualid") & request.Form("cualid")
	if cualid <> "" then
		session("cualid") = cualid
	elseif session("cualid") <> "" then
		cualid = session("cualid")
	else
		unerror = true : msgerror = "No se ha especificado una zona de aSkipper. (Cualidad)"
	end if
	
	if not unerror then
		Response.Redirect("aSkipper.asp")
	else
		Response.Write("<b>Error</b><br>"& msgerror)
	end if
		
end if%>
</body>
</html>
