<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/admin/inc_rutinas.asp" -->
<html>
<head>
<title>Meta-Tags</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
</head>
<body class="bodyAdmin">
<%
	secc = ""& request.QueryString("secc")
	if secc = "" then
		unerror = true : msgerror = "No se ha recibido la sección"
	end if

	dim idioma
	idioma = ""&request.QueryString("idioma")
	archivoXml = "/" & c_s & idioma & secc &"/"& nombreArchivo(secc) & ".xml"

	if not unerror then
		Set xmlObj = CreateObject("MSXML.DOMDocument")
		if not xmlObj.load(server.MapPath(archivoXml)) then
			unerror = true : msgerror = "No se ha encontrado el archivo XML."
		else
			set nodoContenido = xmlObj.selectSingleNode("contenido")
			if not typeOK(nodoContenido) then
				unerror = true : msgerror = "No se ha encontrado el nodo contenido."
			end if
		end if
	end if
	
	if not unerror then
		set ntitulo = nodoContenido.selectSingleNode("titulo")
		if typeOK(ntitulo) then titulo = ntitulo.text end if

		set nkeys = nodoContenido.selectSingleNode("keys")
		if typeOK(nkeys) then keys = nkeys.text end if

		set ndescripcion = nodoContenido.selectSingleNode("descripcion")
		if typeOK(ndescripcion) then descripcion = ndescripcion.text end if
	end if
	
	if not unerror then
	
		if request.Form() <> "" then
			if not typeOK(ntitulo) then
				set ntitulo = xmlObj.createElement("titulo")
				nodoContenido.appendChild(ntitulo)
			end if
			ntitulo.text = filtroHtml(request.Form("titulo"))

			if not typeOK(nkeys) then
				set nkeys = xmlObj.createElement("keys")
				nodoContenido.appendChild(nkeys)
			end if
			nkeys.text = filtroHtml(request.Form("keys"))
			
			if not typeOK(ndescripcion) then
				set ndescripcion = xmlObj.createElement("descripcion")
				nodoContenido.appendChild(ndescripcion)
			end if
			ndescripcion.text = filtroHtml(request.Form("descripcion"))

			xmlObj.save server.MapPath(archivoXml)
			%>
			Un momento ...
			<script language="javascript" type="text/javascript">
				window.close()
			</script>
			<%
		
		else%>
			<table width="100%"  border="0" cellspacing="0" cellpadding="5">
				<tr><td><span class="tituloazonaAdmin">Meta-tags</span><br>
				Las etiquetas &lt;Meta&gt; sirven para complementar informaci&oacute;n que leer&aacute;n los buscadores.<br>
				Desde aqu&iacute; tiene la posibilidad de editar tres de las mas usadas.<br></td></tr>
				<tr><td>
				<form name="form1" method="post" action="metatags.asp?secc=<%=secc%>&idioma=<%=idioma%>">
				T&iacute;tulo:<br>
				<input name="titulo" type="text" class="campoAdmin" style="width:100%" value="<%=titulo%>" maxlength="200">
				<br>
				Palabras clave (Keywords):<br>
				<input name="keys" type="text" class="campoAdmin" style="width:100%" value="<%=keys%>" maxlength="200">
				<br>
				Descripci&oacute;n:<br>
				<textarea name="descripcion" cols="" rows="5" wrap="virtual" class="areaAdmin" style="width:100%"><%=descripcion%></textarea>
				<br>
				<div align="right">
				<input name="" type="button" class="botonAdmin" onClick="window.close()" value="Cancelar">
				<input type="submit" class="botonAdmin" value="Enviar">
				</div>
				<br>
				</form></td></tr>
			</table>
		<%end if
	end if%>


	<%if unerror then
response.Write msgerror
	end if%>	
</body>
</html>