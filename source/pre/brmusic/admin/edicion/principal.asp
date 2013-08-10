<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->

<% 

session("cualid")=""&request.QueryString("cualid")


%>

<html>
<head>
<title>aSkipper</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" class="bodyAdmin">
<%if unerror then
	Response.Write "<b>Error</b><br>" & msgerror
else%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="center">
		<%if not unerror then%>

			<%if session("usuario") <> "" then
				if getPermiso("edicion",session("idioma")) then%>
					<script>
						// Acualizo el marco izquierdo para que me muestra la lista
						parent.frames[0].location.href='lista.asp?<%=request.QueryString()%>'
					</script>
					<b>Bienvenid@</b>
					<%if request("msg") <> "" then%>
						<br>
						<table border="0" cellspacing="0" cellpadding="4">
							<tr>
							<td align="center" bgcolor="#009900"><font color="#FFFFFF"><%=request("msg")%></font></td>
							</tr>
						</table>
					<%end if
				else%>
					<script>
						top.location.href="../usuarios/noacceso.asp"
					</script>
				<%end if
			else%>
			<script>
				top.location.href="../usuarios/nologeado.asp"
			</script>
			<%end if%>
		<%else%>
			<b>Ha ocurrido un error</b><br>
			<%=msgerror%>
		<%end if%>
	</td>
	</tr>
</table>
<%end if%>
</body>
</html>