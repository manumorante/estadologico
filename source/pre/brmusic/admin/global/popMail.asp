<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<!--#include file="inc_seguridad.asp" -->
<!--#include virtual="/admin/inc_rutinas.asp" -->
<!--#include file="inc_rutinas.asp" -->
<!--#include file="inc_conn.asp" -->
<!--#include virtual="/admin/inc_sendmail.asp" -->
<html>
<head>
<title>PopMail</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="estilos.css" rel="stylesheet" type="text/css">
</head>
<body class="bodyAdmin" vlink="#003366" onLoad="window.focus();">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
		<td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Enviar email</font></b></td>
		<td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
	  </tr>
	</table>
<%if request.Form() <> "" then

	id = request.Form("id")
	emailemision = request.Form("emailemision")
	nombreemailemision = request.Form("nombreemailemision")
	emailrecepcion = request.Form("emailrecepcion")
	nombreemailrecepcion = request.Form("nombreemailrecepcion")
	asunto = request.Form("asunto")

	' Partes
	' ------------------------------------------------------------------------------------------------------------------
	HTTP_HOST = request.ServerVariables("HTTP_HOST")
	cuerpo = ""
	cuerpo = cuerpo &"Se ha insertado una nueva orden.<br/>"& vbCrlf
	cuerpo = cuerpo &"Para gestionarla pulse en el siguiente enlace:<br/>"& vbCrlf
	cuerpo = cuerpo &"<a href=http://"& HTTP_HOST &"/"& c_s &"admin/admin.asp?id="& id &">http://"& HTTP_HOST &"/"& c_s &"admin/admin.asp?id="& id &"</a><br/>"& vbCrlf
	cuerpo = cuerpo &"<hr/><br/>"& vbCrlf & vbCrlf
	' ------------------------------------------------------------------------------------------------------------------

	cuerpo = cuerpo &request.Form("cuerpo")
	if sendMail(emailemision, nombreemailemision, emailrecepcion, nombreemailrecepcion, asunto, cuerpo) then%>
		<br/><div align="center"><b>Email enviado.</b></div>
		<script language="javascript" type="text/javascript">
		window.close()
		</script>
	<%else%>
		<br/><div align="center"><b>No se ha podido enviar el email.</b></div>
	<%end if
else
	

	%>
	
	<form name="f" action="popMail.asp" method="post">
	ID: <input name="id" type="text" class="campoAdmin" value="<%=request("id")%>">
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
	  <tr>
		<td colspan="3">Escriba la dirección a la que desea enviar un email y el cuerpo del mismo. </td>
	  </tr>
	  <tr>
	    <td colspan="3"><b>El email se enviará como:</b></td>
      </tr>
	  <tr>
	    <td colspan="3"><input name="nombreemailemision" type="text" class="campoAdmin" id="nombreemailemision" value="<%=getNombreUsuario(session("usuario"))%>" readonly="true">
        <input name="emailemision" type="text" class="campoAdmin" id="emailemision" value="<%=getEmailUsuario(session("usuario"))%>" readonly="true"></td>
      </tr>
	  <tr>
		<td width="33%"><b>Asunto: 
		  
		</b></td>
		<td width="33%"><b>Nombre</b></td>
		<td width="33%"><b>Dirección:</b> </td>
	  </tr>
	  <tr>
		<td>
<%
	asunto = ""& request("asunto")
	if asunto = "" then%>
		<input type="text" value="" name="asunto" class="campoAdmin">
	<%else%>
		<input type="hidden" name="asunto" value="<%=asunto%>">
		<%=asunto%>
	<%end if%>
		</td>
		<td><input name="nombreemailrecepcion" type="text" class="campoAdmin" id="nombreemailrecepcion" style="width:100%" value="<%=request("nombre")%>"></td>
		<td><input name="emailrecepcion" type="text" class="campoAdmin" id="emailrecepcion" style="width:100%" value="<%=request("email")%>"></td>
	  </tr>
	  <tr>
		<td colspan="3"><b>Cuerpo del emal:</b></td>
	  </tr>
	  <tr>
		<td colspan="3"><textarea name="cuerpo" cols="" rows="10" wrap="virtual" class="areaAdmin" id="cuerpo" style="width:100%"><%=request("cuerpo")%></textarea></td>
	  </tr>
	  <tr>
		<td colspan="3" align="right"><input name="" type="button" class="botonAdmin" onClick="window.close()" value="Cerrar">
		<input type="submit" class="botonAdmin" value="Enviar"></td>
	  </tr>
	</table>
	</form>
<%end if%>
</body>
</html>