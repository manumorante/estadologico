<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Documento sin t&iacute;tulo</title>
<link href="/admin/global/estilos.css" rel="stylesheet">
</head>
<body bgcolor="#F5F5F5">
<%select case request.QueryString("ac")%>

<%case "formguardarfotoseccion2"%>

	<form action="archivos.asp?ac=guardarfotoseccion2&id=<%=request.QueryString("id")%>&icono=<%=request.QueryString("icono")%>&foto=<%=request.QueryString("foto")%>" method="post" enctype="multipart/form-data" name="f" id="f" onSubmit="return true;">
		<input name="archivo" type="file" id="archivo" onChange="parent.frames[0].f.archivo.value = this.value;"><br>
		<input name="id" type="hidden" id="id" value="<%=request.QueryString("id")%>">
		<input name="ancho_foto" type="hidden" id="ancho_foto" value="">
		<input name="alto_foto" type="hidden" id="alto_foto" value="">
		<input name="ancho_icono" type="hidden" id="ancho_icono" value="">
		<input name="alto_icono" type="hidden" id="alto_icono" value="">
		<input name="icono" type="hidden" id="icono" value=""><br>
		<input name="enviar" type="submit" id="enviar" value="Enviar" class="boton">
	</form>

<%case "formguardarfotoseccion"%>

	<form action="archivos.asp?ac=guardarfotoseccion&id=<%=request.QueryString("id")%>&icono=<%=request.QueryString("icono")%>&foto=<%=request.QueryString("foto")%>" method="post" enctype="multipart/form-data" name="f" id="f" onSubmit="return true;">
		<input name="archivo" type="file" id="archivo" onChange="parent.frames[0].f.archivo.value = this.value;"><br>
		<input name="id" type="hidden" id="id" value="<%=request.QueryString("id")%>">
		<input name="ancho_foto" type="hidden" id="ancho_foto">
		<input name="alto_foto" type="hidden" id="alto_foto">
		<input name="ancho_icono" type="hidden" id="ancho_icono">
		<input name="alto_icono" type="hidden" id="alto_icono">
		<input name="icono" type="hidden" id="icono"><br>
		<input name="enviar" type="submit" id="enviar" value="Enviar" class="boton">
	</form>

<%case "formguardaricono"%>

	<form action="archivos.asp?ac=guardaricono&id=<%=request.QueryString("id")%>&foto=<%=request.QueryString("foto")%>" method="post" enctype="multipart/form-data" name="f" id="f" onSubmit="return true;">
		<input name="archivo" type="file" id="archivo" onChange="parent.frames[0].f.archivo.value = this.value;"><br>
		<input name="id" type="hidden" id="id" value="<%=request.QueryString("id")%>">
		<input name="ancho_foto" type="hidden" id="ancho_foto">
		<input name="alto_foto" type="hidden" id="alto_foto">
		<input name="ancho_icono" type="hidden" id="ancho_icono">
		<input name="alto_icono" type="hidden" id="alto_icono">
		<input name="icono" type="hidden" id="icono"><br>
		<input name="enviar" type="submit" id="enviar" value="Enviar" class="boton">
	</form>

<%case "formguardarfoto"%>

	<form action="archivos.asp?ac=guardarfoto&id=<%=request.QueryString("id")%>&icono=<%=request.QueryString("icono")%>&foto=<%=request.QueryString("foto")%>" method="post" enctype="multipart/form-data" name="f" id="f" onSubmit="return true;">
		<input name="archivo" type="file" id="archivo" onChange="parent.frames[0].f.archivo.value = this.value;" size="40">
		<br>
		<input name="id" type="hidden" id="id" value="<%=request.QueryString("id")%>">
		<input name="ancho_foto" type="hidden" id="ancho_foto">
		<input name="alto_foto" type="hidden" id="alto_foto">
		<input name="ancho_icono" type="hidden" id="ancho_icono">
		<input name="alto_icono" type="hidden" id="alto_icono">
		<input name="icono" type="hidden" id="icono"><br>
		<input name="enviar" class="botonAdmin" type="submit" id="enviar" value="Enviar">
		<input name="" class="botonAdmin" type="button" id="enviar" onClick="location.href='inicio.asp';" value="Cancelar">
	</form>

<%case else%>
	<!-- <b>Error:</b> Ninguna acción. -->
<%end select%>
</body>
</html>
