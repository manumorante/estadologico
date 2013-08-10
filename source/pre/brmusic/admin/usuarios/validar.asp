<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include file="rutinasParaAdmin.asp" -->
<!--#include virtual="/admin/inc_sha256.asp" -->
<%
'session("usuario") = 1
'session("idioma") = "esp"
'session("zona") = 2

%>
<html>
<head>
<title>Administraci&oacute;n</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript1.1" type="text/javascript">
<!--
	if(window.top.location != window.self.location){
		window.top.location = window.self.location;
	}

//-->
</script>
<link href="../estilos.css" rel="stylesheet" type="text/css">
</head>
<body class="bodyAdmin">

<%
if not unerror then

	select case request.QueryString("ac")
	
	' Cambiar el idioma al instante
	case "cambiaridioma"

		nuevoidioma = request.QueryString("nuevoidioma")
		if nuevoidioma <> "" then
			session("idioma") = nuevoidioma%>
			<script type="text/javascript">top.location.href="../aSkipper.asp"</script>
		<%end if
	
	
	case "desconectar"
		
	session("usuario") = ""
	session("idioma") = ""
	session("zona") = ""
	session.Abandon()
	Response.Redirect("validar.asp")
	
	case else
	
		if session("usuario") <> "" and session("idioma") <> "" then
			Response.Redirect("../aSkipper.asp?"& request.QueryString())
		end if

	
		if request.Form() <> "" then
			dim c_usuario, c_clave
			c_usuario = request.Form("usuario")
			c_clave = SHA256(request.Form("clave"))
			if c_usuario <> "" and c_clave <> "" then
				if c_clave = getClave(c_usuario) then
					' Declaro la sesion de usuario con su código
					session("usuario") = getCodigo(c_usuario)
'					session("idioma") = getIdioma(c_usuario)
					session("idioma") = "esp" ' Español por defecto.
					session("zona") = 2
		
					Response.Redirect("../aSkipper.asp")
				else
					errorvalidar = true : msgerrorvalidar = "Usuario o clave incorrectos."
				end if
			else
				errorvalidar = true : msgerrorvalidar = "Debe escribir su nombre y clave de usuario."
			end if
		end if
		
		if request.Form() = "" or errorvalidar then%>
		<br>
		<form action="validar.asp" method="post" name="f" id="f1">
		  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td align="center" valign="middle"><table width="100" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
                <tr align="center">
                  <td colspan="2"><strong><font color="#000000" size="3">Administraci&oacute;n</font></strong></td>
                </tr>
                <tr>
                  <td align="right"><span class="Estilo1">Usuario</span></td>
                  <td><input name="usuario" type="text" class="campo" id="usuario"></td>
                </tr>
                <tr>
                  <td align="right"><span class="Estilo1">Clave</span></td>
                  <td><input name="clave" type="password" class="campo" id="clave"></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="right">&nbsp;</td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td align="right"><input type="button" class="boton" onClick="location.href='../../'" value="Salir">
                      <input type="submit" class="boton" value="Aceptar"></td>
                </tr>
              </table>
              <%if request.QueryString("msg") <> "" or msgerrorvalidar <> "" then%>
                                  <br>
                                  <table  border="0" cellpadding="2" cellspacing="0">
                                    <tr>
                                      <td bgcolor="#f5f5f5"><font color="#990000"><%=request.QueryString("msg")&msgerrorvalidar%></font></td>
                                    </tr>
                </table>
                
					  <%end if%></td>
            </tr>
          </table>
		  <script>f.usuario.focus()</script>
		</form>
		<%end if%>
	<%end select%>

<%else  ' de [unerror] principal%>
	Ha ocurrido un error.<br>
	<b><%=msgerror%></b>
<%end if ' de [unerror] principal%>
</body>
</html>
