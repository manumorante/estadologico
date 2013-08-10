<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Skipper</title>
</head>
<body>
<script language="javascript" type="text/javascript">
	<!--
	function reciboRuta() {
		top.frames[0].f.fotoNueva.value = f.archivo.value
	}
	//-->
</script>
<form action="imagen.asp?ac=ok&num=<%=request("num")%>&secc=<%=request("secc")%>&idi=<%=request("idi")%>" method="post" enctype="multipart/form-data" name="f">
<input name="archivo" type="file" id="archivo" onChange="reciboRuta();">
<br> 
Archivo actual
:
<input name="archivo_actual" type="text" id="archivo_actual">

Ancho:
<input name="ancho" type="text" id="ancho">

Alto: 
<input name="alto" type="text" id="alto">
<br>
Pie:
<input name="pie" type="text" id="pie"> 
Enlace:
<input name="enlace" type="text" id="enlace">

Enlace ventana:
<input name="enlaceventana" type="text" id="enlaceventana">

novisible:
<input name="novisible" type="checkbox" id="novisible" value="1">
<br>
margen:
<input name="margen" type="checkbox" id="margen" value="1">

Editar: 
<input name="editar" type="checkbox" id="editar" value="1">
<br>

<input name="enviar" type="submit" id="enviar" value="Enviar">
</form>
</body>
</html>
