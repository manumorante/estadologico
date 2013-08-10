<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/inc_sha256.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<%
	idioma = ""&request.QueryString("idioma")
	if idioma = "" then
		unerror = true : msgerror = "No ha sido posible determinar el idioma de navegación / aSkipper"
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
<title>aSkipper - Entrar</title>
<script language="javascript" type="text/javascript">
<!--
function cerrar(){
window.close()
}
//-->
</script>

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:8.5pt;
}
.Rojo {color: #FF0000}
.Verde {color: #009900}
.Estilo2 {
	color: #006600;
	font-weight: bold;
}
.Estilo4 {color: #FF0000; font-weight: bold; }
-->
</style></head>
<body bgcolor="#F9F9F9"><img src="../images/logo_Skipper.jpg" width="273" height="35">
<%if not unerror then
	if ""&session("usuario") <> "" and ""&session("zona") <> "2" then
		if getPermiso("edicion",idioma) then
			session("zona") = 2
		end if
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><br>
<br>

		<div align="center"><span class="Estilo2">Activando</span><br>
		  Idioma: <b><%=getNombreIdioma(idioma)%></b></div></td>
          </tr>
</table>
		

		<script language="javascript" type="text/javascript">
		<!--
			parent.opener.location = parent.opener.location
			setTimeout("cerrar()",800)
		//-->
		</script>
		<%
	elseif ""&session("usuario") <> "" and ""&session("zona") = "2" then
		session("zona") = 0
		%>
		<table width="100%" border="0" cellpadding="5" cellspacing="0">
          <tr>
            <td><br>
                <br>
            <div align="center"><span class="Estilo4">Desactivando</span><br>
            Idioma: <b><%=getNombreIdioma(idioma)%></b></div></td>
          </tr>
        </table>
		<script language="javascript" type="text/javascript">
		<!--
			parent.opener.location = parent.opener.location
			setTimeout("cerrar()",800)
		//-->
		</script>
		<%
	else
		c_usuario = ""&request.Form("usuario")
		c_clave = sha256(""&request.Form("clave"))
		if c_usuario <> "" and c_clave <> "" then
			if c_clave = getClave(c_usuario) then
				session("usuario") = getCodigo(c_usuario)
				session("idioma") = idioma
				if getPermisoPara("edicion","",session("usuario")) then
					session("zona") = 2
				end if
				%><br><br>
				<div align="center">Bienvenid@ <font color="#006600"><b><%=c_usuario%></b></font><br>
				Idioma: <b><%=getNombreIdioma(idioma)%></b></div>
				<script language="javascript" type="text/javascript">
				<!--
					parent.opener.location = parent.opener.location
					setTimeout("cerrar()",1000)
				//-->
				</script>
				<%
			else
				%>
				<div align="center"><br>

			      <br>
			      <b>Usuario o clave incorrectos</b><br>
			    <br>
			    <input name="" type="button" class="botonAdmin" onClick="location.href='entrar.asp?idioma=<%=idioma%>'" value="Reintentar">
				</div>
				<%
			end if

		else
		%><form name="f" method="post" action="entrar.asp?idioma=<%=idioma%>">
		  <table  border="0" align="center" cellpadding="4" cellspacing="0">
            <tr>
              <td align="right">Usuario</td>
              <td><input name="usuario" type="text" class="campoAdmin" id="usuario"></td>
            </tr>
            <tr>
              <td align="right">Clave</td>
              <td><input name="clave" type="password" class="campoAdmin" id="clave"></td>
            </tr>
          </table>
		  <table width="100%"  border="0" align="center" cellpadding="4" cellspacing="0">
            <tr>
              <td align="right"><input name="Enviar" type="submit" class="botonAdmin" value="Entrar"></td>
            </tr>
          </table>
		  <script language="javascript" type="text/javascript">f.usuario.focus()</script>
</form>
		<%
		end if
	end if
end if

if unerror then
	Response.Write msgerror
end if
%>
</body>
</html>
