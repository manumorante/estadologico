<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Privado</title>
<link href="estilos.css" rel="stylesheet" type="text/css">
</head>

<body>
<p><b>Privado
</b></p>
<%
if ""& request.Form("usuario") = "admin" and ""& request.Form("clave") = "Qlvytden5" then
	session("usuario") = 1
	Response.Redirect("index.asp")
elseif ""& request.Form("usuario") = "invitado" and ""& request.Form("clave") = "ver" then
	session("usuario") = 2
	Response.Redirect("index.asp")
end if
%>
<form action="#" method="post" name="f" id="f">
  <table  border="0" align="center" cellpadding="4" cellspacing="0">
    <tr>
      <td align="right">Usuario</td>
      <td><input name="usuario" type="text" class="campo" id="usuario"></td>
    </tr>
    <tr>
      <td align="right">Clave</td>
      <td><input name="clave" type="password" class="campo" id="clave"></td>
    </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td align="right"><input type="submit" value="Enviar"></td>
    </tr>
  </table>
</form>
<p>&nbsp; </p>
</body>
</html>
