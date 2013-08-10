<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' Evito que salga la cabecera de usa
zona=1
%>
<!--#include virtual="/admin/usa.asp" -->
<%
dim ruta

%>
<html>
<head>
<title>Administraci&oacute;n</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
<link href="../../estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo3 {color: #006600}
a:link {
	text-decoration: underline;
}
a:visited {
	text-decoration: underline;
}
a:hover {
	text-decoration: underline;
}
a:active {
	text-decoration: underline;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" class="bodyAdmin">
<%
select case request.QueryString("ac")
case "editar"


' Comprobación de variables recibidas
if isNumeric(request.QueryString("fila")) and isNumeric(request.QueryString("columna")) and isNumeric(request.QueryString("numTabla")) and ""&secc <> "" then
	Dim fila
	dim columna
'	dim texto
	numTabla = request.QueryString("numTabla")
	fila = request.QueryString("fila")
	columna = request.QueryString("columna")
	abs_xml = "/"& c_s & session("idioma")& secc & "/"& nombreArchivo(secc) & ".xml"

	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing" then
			eltexto = latabla.childNodes.item(fila).childNodes.item(columna).text
		else ' Error, no se encuentra el nodo indicado.
			unerror = true
			msgerror = "<br><b>Error, no se encuentra el nodo indicado.</b>"
		end if
	else ' Error, no se enuentra el archivo xml.
		unerror = true
		msgerror = "<br><b>Error, no se encuentra el archivo xml.</b>"
	end if
else ' Error, faltan datos.
	unerror = true
	msgerror = "<br>Error, faltan datos."
end if

If unerror then
	%><script>
	var elmsgerror
	elmsgerror = "<%=msgerror%>"
	elmsgerror = elmsgerror.replace(/<b>|<\/b>|<br>/g,'')
	alert(elmsgerror)
	window.history.back()
	</script><%
end if



%>

<script language="Javascript1.2">
<!--
// -------------------------------------------------------------------------------  htmlarea
_editor_url = "";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
	document.write('<scr' + 'ipt src="' +_editor_url+ 'Editor.js"');
	document.write(' language="Javascript1.2"></scr' + 'ipt>');
} else {
	document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>');
}
// No declaro misBotones para que coje la botonera por defecto
var	misBotones = 0
//-------------------------------------------------------------------------------------------------------- -->
</script>

<form name="f" action="tabla.asp?ac=editar_ya" method="post">
<input name="fila" type="hidden" value="<%=fila%>">
	<input name="columna" type="hidden" value="<%=columna%>">
	<input name="numTabla" type="hidden" value="<%=numTabla%>">
	<input name="secc" type="hidden" value="<%=secc%>">
<textarea name="texto" cols="" rows="" wrap="virtual" class="areaAdmin" id="texto" style="WIDTH=100%;HEIGHT=88%"><%=Replace(eltexto,"<br>",vbCrLf)%></textarea>
	<script language="javascript1.2">
editor_generate('texto');
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr valign="bottom">
    <td height="25"  align="left">        &nbsp;
      <input name="" type="button" class="botonAdmin" title=" Borrar el contenido de esta celda " onClick="texto.value='&nbsp;';f.submit()" value="Dejar celda vacia">        </td>
    <td height="25"  align="right"><input name="Submit" type="submit" class="botonAdmin" value="Enviar">
      <input type="button" class="botonAdmin" onClick="window.close()" value="Cancelar">
      &nbsp;</td>
  </tr>
</table>
</form>

<%
case "editar_ya"
unerror = false : msgerror = ""
if trim(request.Form("fila")) <> "" and trim(request.Form("columna")) <> "" and trim(request.Form("numTabla")) <> "" then
	fila = trim(request.Form("fila"))
	columna = trim(request.Form("columna"))
	numTabla = trim(request.Form("numTabla"))
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if trim(request.Form("texto")) <> "" then
	eltexto = trim(request.Form("texto"))
elseif not unerror then
	unerror = true
	msgerror = "Atención:\n\nSi desea dejar la celda vacía, por favor, utilice el botón 'Dejar celda vacia'.\nEste botón inserta en espacio en blanco para causar el efecto deseado."
end if

if ""&request.Form("secc") <> "" then
	secc = ""&request.Form("secc")
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	abs_xml = "/"& c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			eltexto = replace(eltexto,"€","&euro;")
			eltexto = replace(eltexto,"·","&#8226;")
			eltexto = Replace(eltexto,vbCrLf,"<br>")
			latabla.childNodes.item(fila).childNodes.item(columna).text = filtroHtml(eltexto)
			xmlObj.save Server.MapPath(abs_xml)
			%>
			<script>
				parent.opener.location.href=parent.opener.location
				window.close()
			</script>
			<%
'			Response.Redirect("tabla.asp?ac=editar&fila="&fila&"&columna="&columna&"&secc"&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML.<br>["&abs_xml&"]"
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if

case "nuevacolumna"

unerror = false : msgerror = ""
if isNumeric(trim(request.QueryString("columna"))) and isNumeric(trim(request.QueryString("numTabla"))) then
	columna = trim(request.QueryString("columna"))
	numTabla = trim(request.QueryString("numTabla"))
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			Dim nuevaColumna
			for n=0 to latabla.childNodes.length-1
				Set nuevaColumna = xmlObj.createElement("columna")
				nuevaColumna.text = ""
				latabla.childNodes.item(n).insertBefore nuevaColumna, latabla.childNodes.item(n).childNodes.item(columna)
			next
			xmlObj.save Server.MapPath(abs_xml)
			Set nuevaColumna = Nothing
			Set xmlObj = nothing
			Set latabla = nothing
			
			
			
			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if


case "resaltarcolumna"

unerror = false : msgerror = ""
if isNumeric(trim(request.QueryString("columna"))) and isNumeric(trim(request.QueryString("numTabla"))) then
	columna = trim(request.QueryString("columna"))
	numTabla = trim(request.QueryString("numTabla"))
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			Dim resalte, valorResalte
			for n=0 to latabla.childNodes.length-1
				if latabla.childNodes.item(n).childNodes.item(columna-1).getAttribute("resalte") = "1" then
					valorResalte = "0"
				else
					valorResalte = "1"
				end if
				Set resalte = xmlObj.createAttribute("resalte")
				latabla.childNodes.item(n).childNodes.item(columna-1).setAttributeNode(resalte)
				resalte.nodeValue= valorResalte
			next
			xmlObj.save Server.MapPath(abs_xml)
			Set xmlObj = nothing
			Set latabla = nothing

			

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if


case "alinearcolumna"

unerror = false : msgerror = ""
if isNumeric(trim(request.QueryString("columna"))) and isNumeric(trim(request.QueryString("dir"))) and isNumeric(trim(request.QueryString("numTabla"))) then
	Dim dir
	dir = trim(request.QueryString("dir"))
	columna = trim(request.QueryString("columna"))
	numTabla = trim(request.QueryString("numTabla"))
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			Dim atribAlin,alineado
			for n=0 to latabla.childNodes.length-1
				select case dir
				case 0
					alineado = ""
				case 1
					alineado = "left"
				case 2
					alineado = "center"
				case 3
					alineado = "right"
				case else
					alineado = "left"
				end select
				Set atribAlin = xmlObj.createAttribute("alineado")
				latabla.childNodes.item(n).childNodes.item(columna-1).setAttributeNode(atribAlin)
				atribAlin.nodeValue= alineado
			next
			xmlObj.save Server.MapPath(abs_xml)
			Set xmlObj = nothing
			Set latabla = nothing

						

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if


' ---------------------------------------------------------------------------------------------
case "borrarcolumna"

unerror = false : msgerror = ""
if isNumeric(trim(request.QueryString("columna"))) and isNumeric(trim(request.QueryString("numTabla"))) then
	columna = trim(request.QueryString("columna"))
	numTabla = trim(request.QueryString("numTabla"))
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			for n=0 to latabla.childNodes.length-1
				latabla.childNodes.item(n).removeChild(latabla.childNodes.item(n).childNodes.item(columna))
			next
			xmlObj.save Server.MapPath(abs_xml)
			Set xmlObj = nothing
			Set latabla = nothing

			

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if



case "nuevafila" ' -----------------------------------------------------------------------------------------------

Dim num
unerror = false : msgerror = ""
if isNumeric(trim(request.QueryString("fila"))) and isNumeric(trim(request.QueryString("numTabla"))) and isNumeric(trim(request.QueryString("num"))) then
	num = trim(request.QueryString("num"))
	fila = trim(request.QueryString("fila"))
	numTabla = trim(request.QueryString("numTabla"))
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then 
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			Dim nuevaFila
			Set nuevaFila = xmlObj.createElement("fila")
			for n=0 to num-1
				miResalte = cint(0 & latabla.childNodes.item(fila).childNodes.item(n).getAttribute("resalte"))
				Set nuevaColumna = xmlObj.createElement("columna")
				set attResalte = xmlObj.createAttribute("resalte")
				nuevaColumna.setAttributeNode(attResalte)
				attResalte.nodeValue = miResalte
				nuevaFila.appendChild(nuevaColumna)
				nuevaColumna.text = ""
				set attResalte = nothing
			next
			latabla.insertBefore nuevaFila, latabla.childNodes.item(fila+1)
			xmlObj.save Server.MapPath(abs_xml)
			Set nuevaFila = Nothing
			Set xmlObj = nothing
			Set latabla = nothing
			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if
' ------------------------------------------------------------------------------------------------------

case "resaltarFila" ' -----------------------------------------------------------------------------------------------

unerror = false : msgerror = ""
if isNumeric(request.QueryString("fila")) and isNumeric(request.QueryString("numTabla")) and isNumeric(request.QueryString("num")) then
	num = request.QueryString("num")
	fila = request.QueryString("fila")
	numTabla = request.QueryString("numTabla")
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing" then
			if latabla.childNodes.item(fila).getAttribute("resalte") = "1" then
				valorResalte = "0"
			else
				valorResalte = "1"
			end if
			set resalte = xmlObj.createAttribute("resalte")
			latabla.childNodes.item(fila).setAttributeNode(resalte)
			resalte.nodeValue = valorResalte
			xmlObj.save Server.MapPath(abs_xml)
			Set xmlObj = nothing
			Set latabla = nothing
			
			

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if
' ------------------------------------------------------------------------------------------------------

case "alinearFila" ' -----------------------------------------------------------------------------------------------

unerror = false : msgerror = ""
if isNumeric(request.QueryString("fila")) and isNumeric(request.QueryString("dir")) and isNumeric(request.QueryString("numTabla")) and isNumeric(request.QueryString("num")) then
	dir = request.QueryString("dir")
	num = request.QueryString("num")
	fila = request.QueryString("fila")
	numTabla = request.QueryString("numTabla")
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing" then
				select case dir
				case 0
					alineado = ""
				case 1
					alineado = "left"
				case 2
					alineado = "center"
				case 3
					alineado = "right"
				case else
					alineado = "left"
				end select
			set atribAlin = xmlObj.createAttribute("alineado")
			latabla.childNodes.item(fila).setAttributeNode(atribAlin)
			atribAlin.nodeValue = alineado
			xmlObj.save Server.MapPath(abs_xml)
			Set xmlObj = nothing
			Set latabla = nothing

			

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if
' ------------------------------------------------------------------------------------------------------



case "borrarfila" ' -----------------------------------------------------------------------------------------------

unerror = false : msgerror = ""
if isNumeric(trim(request.QueryString("fila"))) and isNumeric(trim(request.QueryString("numTabla"))) then
	fila = trim(request.QueryString("fila"))
	if fila <0 then fila = 0 end if
	numTabla = trim(request.QueryString("numTabla"))
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			latabla.removeChild(latabla.childNodes.item(fila))
			xmlObj.save Server.MapPath(abs_xml)
			Set xmlObj = nothing
			Set latabla = nothing

			

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if
' ------------------------------------------------------------------------------------------------------

case "nuevatabla"

unerror = false : msgerror = ""
if isNumeric(trim(request.QueryString("filas"))) and isNumeric(trim(request.QueryString("columnas"))) and isNumeric(trim(request.QueryString("numTabla"))) then
	dim filas, columnas
	filas = trim(request.QueryString("filas"))
	columnas = trim(request.QueryString("columnas"))
	numTabla = trim(request.QueryString("numTabla"))
else
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido la celda para editar."
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") & "/" & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			for n2 = 0 to filas-1
				Set nuevaFila = xmlObj.createElement("fila")
				if columnas = 0 then columnas = 1 end if
				for n=0 to columnas-1
					Set nuevaColumna = xmlObj.createElement("columna")
					nuevaFila.appendChild(nuevaColumna)
					nuevaColumna.text = " "
				next
				latabla.insertBefore nuevaFila, latabla.childNodes.item(fila+1)
			next
			xmlObj.save Server.MapPath(abs_xml)
			Set nuevaFila = Nothing
			Set xmlObj = nothing
			Set latabla = nothing

			

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if
' ------------------------------------------------------------------------------------------------------

case "dimensiones"

unerror = false : msgerror = ""
if request.Form("numTabla") <> "" then
	dim ancho, alto
	ancho = request.Form("ancho")
	alto = request.Form("alto")
	numTabla = request.Form("numTabla")
else
	unerror = true
	msgerror = "<b>Error</b>, faltan datos."
end if

if ""&request.Form("secc") <> "" then
	secc = ""&request.Form("secc")
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			Dim elancho, elalto
			set elancho = xmlObj.createAttribute("ancho")
			latabla.setAttributeNode(elancho)
 			elancho.nodeValue = ancho
			set elalto = xmlObj.createAttribute("alto")
			latabla.setAttributeNode(elalto)
			elalto.nodeValue = alto
			xmlObj.save Server.MapPath(abs_xml)
			Set nuevaFila = Nothing
			Set xmlObj = nothing
			Set latabla = nothing

			

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if


If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if


' ************************************************************************************************ TABLA VISIBLE
case "tablavisible"

unerror = false : msgerror = ""
if request.QueryString("estado") <> "" and request.QueryString("numTabla") <> "" then
	dim estado
	estado = request.QueryString("estado")
	select case request.QueryString("estado")
	case "true" estado = 1
	case "false" estado = 0
	end select
	numTabla = request.QueryString("numTabla")
else
	unerror = true
	msgerror = "<b>Error</b>, faltan datos. (E.1)"
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/" & c_s & session("idioma") &  secc & "/" & nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			Dim visible
			set visible = xmlObj.createAttribute("visible")
			latabla.setAttributeNode(visible)
			visible.nodeValue = estado

			xmlObj.save Server.MapPath(abs_xml)
			Set xmlObj = nothing
			Set latabla = nothing

			

			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if

If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if

case "comentario" ' ---------------------------------------------------------------------------------------------------

unerror = false : msgerror = ""
if request.form("comentario") <> "" and request.QueryString("numTabla") <> "" then
	DIM elcomentario : elcomentario = request.form("comentario")
	numTabla = request.QueryString("numTabla")
else
	unerror = true
	msgerror = "<b>Error</b>, faltan datos. (E.1)"
end if

if ""&secc <> "" then
	secc = ""&secc
elseif not unerror then
	unerror = true
	msgerror = "<b>Error</b>, no se ha recibido el nombre de fichero para editar."
end if

if not unerror then
	abs_xml = "/"& c_s & session("idioma") & secc & "/"& nombreArchivo(secc) & ".xml"
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(abs_xml)) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing"then
			Dim comentario
			set comentario = xmlObj.createAttribute("comentario")
			latabla.setAttributeNode(comentario)
			comentario.nodeValue = elcomentario

			xmlObj.save Server.MapPath(abs_xml)
			Set xmlObj = nothing
			Set latabla = nothing


			Response.Redirect("tabla.asp?secc="&secc&"&numTabla="&numTabla)
		else
			unerror = true
			msgerror = "<b>Error</b>, no se encuentra el nodo indicado."
		end if
	else
		Set xmlObj = nothing
		unerror = true
		msgerror = "<b>Error</b>, no se encuentra el archivo XML."
	end if
end if

If unerror then
	%><script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script><%
end if

case else

if ""&secc <> "" and ""&request.QueryString("numTabla")<>"" then
'	Dim tabla
	dim secc
	dim numTabla
	dim latabla
	numTabla = request.QueryString("numTabla")

	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath("/" & c_s & session("idioma") & secc & "/" & nombreArchivo (secc)&".xml")) then
		set latabla = xmlObj.selectSingleNode("contenido/tabla"&numTabla)
		if typename(latabla) <> "Nothing" then
				visible = latabla.getAttribute("visible")
				comentario = latabla.getAttribute("comentario")

				' Estilos:
				' general-tabla-ext 	La rabla exterior
				' general-tabla 		La tabla interior
				' general-tabla-titulo 	La primeta fila
				' general-tabla-celdas1 Primer color de fila
				' general-tabla-celdas1 Segundo color de fila
				Dim color_n, d, numColumnas
				set filas = latabla.childNodes
				color_n = true
				Dim n, n2, color1, color2, colorCelda
				%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="40" bgcolor="F0F5F5"> <img src="img_admin/logo.gif" width="130" height="39"> 
                </td>
              </tr>
              <tr> 
                <td><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" background="img_admin/titulo_f.gif">
                    <tr> 
                      <td width="5"><font size="3" face="Georgia, Times New Roman, Times, serif"><img src="img_admin/titulo_i.gif" width="4" height="25"></font></td>
                      <td><span class="general-ladillo"><%=comentario%></span></td>
                      <td width="5" align="right"><img src="img_admin/titulo_d.gif" width="5" height="25"></td>
                    </tr>
                  </table></td>
              </tr>
</table>
<script language="javascript" type="text/javascript">
function ventana (laURL,ancho,alto) {
	var winl = (screen.width - ancho) / 2;
	var wint = (screen.height - alto) / 2;
	var paramet='top='+wint+',left='+winl+',width='+ancho+',height='+alto;
	var splashWin=window.open(laURL,'Admin',paramet);
    splashWin.focus();
}
	function editarCelda(fila,columna) {
		var laURL = "tabla.asp?ac=editar&fila="+fila+"&columna="+columna+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
		ventana(laURL,650,450)
	}
// --------
	function nuevaColumna(columna) {
		location.href="tabla.asp?ac=nuevacolumna&columna="+columna+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
	}
	function alinearColumna(columna,dir) {
		location.href="tabla.asp?ac=alinearcolumna&columna="+columna+"&dir="+dir+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
	}
	function resaltarColumna(columna) {
		location.href="tabla.asp?ac=resaltarcolumna&columna="+columna+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
	}
	function borrarColumna(columna,numColumnas) {
		if (numColumnas > 1) {
			if (confirm("¿Está seguro que quiere borrar la columna y todos sus contenidos?")){
				location.href="tabla.asp?ac=borrarcolumna&columna="+columna+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
			}
		} else {
			alert("No puede borrar la última columna.")
		}
	}
	function nuevaFila(fila,num) {
		location.href="tabla.asp?ac=nuevafila&num="+num+"&fila="+fila+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
	}
	function resaltarFila(fila,num) {
		location.href="tabla.asp?ac=resaltarFila&num="+num+"&fila="+fila+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
	}
	function alinearFila(fila,num,dir) {
		location.href="tabla.asp?ac=alinearFila&num="+num+"&fila="+fila+"&dir="+dir+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
	}
	function borrarFila(fila,num) {
		if(num>1) {
			if (confirm("¿Está seguro que quiere borrar la fila y todos sus contenidos?")){
				location.href="tabla.asp?ac=borrarfila&fila="+fila+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
			}
		}else{
			alert("No puede borrar la última fila.")
		}
	}
	function nuevatabla(filas,columnas) {
		location.href="tabla.asp?ac=nuevatabla&filas="+filas+"&columnas="+columnas+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
	}

				function validarDimensiones(f) {
					return true
				}
				function tablaVisible(estado) {
					location.href="tabla.asp?ac=tablavisible&estado="+estado+"&secc=<%=secc%>&numTabla=<%=numTabla%>"
				}
				</script> 
<br>
<table width="98%"  border="0" align="center" cellpadding="2" cellspacing="0">
  <tr>
    <td><span class="tituloazonaAdmin">Edici&oacute;n de tabla</span> <br>
      Edite el contenido de las celdas pulsando en el icono      <img src="/<%=c_s%>admin/global/img/lapiz.gif" alt=" Editar tabla " border="0">      .<br>
        Puede insertar nuevas filas y columnas o eliminar existentes pulsando en los iconos.<br>
      Para ver la funcionalidad de los iconos sit&uacute;e el cursor encima y espere un segundo. </td>
  </tr>
</table>
<br>
<form action="tabla.asp?ac=dimensiones" method="post" name="dimensiones" onSubmit="return validarDimensiones(this)">
<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
	
<fieldset><legend>Propiedades de la tabla</legend>
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="0">
<tr> 
                <td align="center" valign="middle">
                    <table border="0" cellpadding="1" cellspacing="0">
                      <tr> 
                        <td colspan="3" align="right">
						<nobr>Ancho: <input name="ancho" type="text" class="campoAdmin" id="ancho" value="<%=latabla.getAttribute("ancho")%>" size="5" maxlength="4"></nobr>
						<nobr>Alto: <input name="alto" type="text" class="campoAdmin" id="alto" value="<%=latabla.getAttribute("alto")%>" size="5" maxlength="4"></nobr>
						<nobr><input name="Enviar" type="submit" class="botonAdmin" title=" Aplicar los valores introducidos en los campos Ancho y Alto " value="Aplicar">
						<input name="" type="button" class="botonAdmin" title=" Aplicar valores predefinidos " onClick="ancho.value = &quot;100%&quot;;alto.value=&quot; &quot;; dimensiones.submit()" value="Reset"></nobr>
						<input name="numTabla" type="hidden" id="numTabla" value="<%=numTabla%>">
						<input name="secc" type="hidden" id="secc" value="<%=secc%>"></td>
                      </tr>

                  </table></td>
                <td align="center" valign="middle">
				  <table border="0" cellpadding="0" cellspacing="0">
				<tr> 
				<td align="center"><%if Cstr(""&visible)="1" then%>
					Esta tabla es <font color="#006600">visible</font>. <input name="" type="button" class="botonAdmin" onClick="tablaVisible(0)" value="Ocultar">
				<%else%>
				Esta tabla est&aacute; <font color="#ff0000">oculta</font>. 
				<input name="" type="button" class="botonAdmin" onClick="tablaVisible(1)" value="Mostrar">
				<%end if%></td>
                  </tr>
				  </table>
		  </td>
</tr>
	  </table>
</fieldset>
	
	</td>
  </tr>
</table>
</form>
<table width="98%"  border="0" align="center" cellpadding="2" cellspacing="0">
  <tr>
                <td><span class="tituloazonaAdmin">                Recuerde</span> <br>
                  La apariencia de esta tabla puede variar en peque&ntilde;os detalles respecto a la original, debido a el espacio que ocupan los iconos de edici&oacute;n.<br>
                  <br></td>
  </tr>
            </table>
            <table class="general-tabla-resaltado-extra" width="<%=latabla.getAttribute("ancho")%>" height="<%=latabla.getAttribute("alto")%>">
              <tr>
                <td> <table width="100%" height="100%" class="general-tabla" cellspacing="2">
                    <%
				dim datos,claseCelda : datos = true
				if typename(latabla.childNodes.item(0)) <> "Nothing" then
					if typename(latabla.childNodes.item(0).childNodes.item(0)) <> "Nothing" then
						for n=0 to filas.length-1
							if color_n then
								color_n = false
								colorCelda = "general-tabla-celdas1"
							else
								color_n = true
								colorCelda = "general-tabla-celdas2"
							end if
							%>
                    <tr>
                      <%
							set columnas = filas.item(n).childNodes
							for n2=0 to columnas.length-1
								numColumnas = getMayor(numColumnas,columnas.length)
								if columnas.item(n2).getAttribute("resalte") = "1" or  filas.item(n).getAttribute("resalte") = "1" then
									claseCelda = "general-tabla-titulo"
								else
									claseCelda = colorCelda
								end if
								' Alineado
								Dim alinFila, alinColumna
								if filas.item(n).getAttribute("alineado") <> "" then
									alinFila = filas.item(n).getAttribute("alineado")
								else
									alinFila = ""
								end if
								if columnas.item(n2).getAttribute("alineado") <> "" then
									alinColumna = columnas.item(n2).getAttribute("alineado")
								else
									alinColumna = ""
								end if
								if alinFila <> "" then
									alineado = alinFila
								elseif alinColumna <> "" then
									alineado = alinColumna
								else
									alineado = ""
								end if

									%>
                      <td class='<%=claseCelda%>' align="<%=alineado%>" valign="top"><a href="JavaScript:editarCelda(<%=n%>,<%=n2%>)"><img src="/<%=c_s%>admin/global/img/lapiz.gif" alt=" Editar celda " border="0"></a>
                      <%=filas.item(n).childNodes.item(n2).text%></td><%
									if n2=columnas.length-1 then
										%>
                      <td valign="top" bgcolor="#FBFBFB"><table border="0" cellspacing="0" cellpadding="0">
                        <tr align="center">
                          <td><a href="#" onClick="nuevaFila(<%=n%>,<%=numColumnas%>)"><img src="img_tabla/nuevaFila.gif" alt="Inserta fila" width="22" height="19" border="0"></a></td>
                          <td><a href="#" onClick="borrarFila(<%=n%>,<%=filas.length%>)"><img src="img_tabla/borrarFila.gif" alt="Eliminar fila" width="22" height="19" border="0"></a></td>
                          <td><a href="#" onClick="resaltarFila(<%=n%>,<%=numColumnas%>)"><img src="img_tabla/resalteFila.gif" alt="Resaltar fila" width="22" height="19" border="0"></a></td>
                        </tr>
                        <tr align="center">
                          <td><%if alinFila = "left" then%>
                        <a href="#" onClick="alinearFila(<%=n%>,<%=numColumnas%>,0)"><img src="img_tabla/alinearIzqAct.gif" alt="La celda está alineada a la izquierda" width="22" height="19" border="0"></a> 
                        <%else%>
                        <a href="#" onClick="alinearFila(<%=n%>,<%=numColumnas%>,1)"><img src="img_tabla/alinearIzq.gif" alt="Alinear celdas a la izquierda" width="22" height="19" border="0"></a> 
                        <%end if%></td>
                          <td><%if alinFila = "center" then%>
                        <a href="#" onClick="alinearFila(<%=n%>,<%=numColumnas%>,0)"><img src="img_tabla/centrarAct.gif" alt="Centrar celdas" width="22" height="19" border="0"></a> 
                        <%else%>
                        <a href="#" onClick="alinearFila(<%=n%>,<%=numColumnas%>,2)"><img src="img_tabla/centrar.gif" alt="Centrar celdas" width="22" height="19" border="0"></a> 
                        <%end if%></td>
                          <td><%if alinFila = "right" then%>
                        <a href="#" onClick="alinearFila(<%=n%>,<%=numColumnas%>,0)"><img src="img_tabla/alinearDerAct.gif" alt="Alinear celdas a la derecha" width="22" height="19" border="0"></a> 
                        <%else%>
                        <a href="#" onClick="alinearFila(<%=n%>,<%=numColumnas%>,3)"><img src="img_tabla/alinearDer.gif" alt="Alinear celdas a la derecha" width="22" height="19" border="0"></a> 
                        <%end if%></td>
                        </tr>
                      </table>
                      </td>
                      <%
								end if
							next
							%>
                    </tr>
                    <%
						next
						%>
                    <tr>
                      <%
						for d=1 to numColumnas
							%>
                      <td bgcolor="#FBFBFB" align="right" valign="top">
					  <table border="0" cellspacing="0" cellpadding="0">
  <tr align="center">
    <td><a href="#"><img src="img_tabla/resalteColumna.gif" alt="Resaltar columna" width="22" height="19" border="0" onClick="resaltarColumna(<%=d%>)"></a></td>
    <td><a href="#"><img src="img_tabla/borrarColumna.gif" alt="Eliminar columna" width="22" height="19" border="0" onClick="borrarColumna(<%=d-1%>,<%=numColumnas%>)"></a></td>
    <td><a href="#"><img src="img_tabla/nuevaColumna.gif" alt="Insertar columna" width="22" height="19" border="0" onClick="nuevaColumna(<%=d%>)"></a></td>
  </tr>
  <tr align="center">
    <td><a href="#" onClick="alinearColumna(<%=d%>,1)"><img src="img_tabla/alinearIzq.gif" alt="Alinear celdas a la izquierda" width="22" height="19" border="0"></a></td>
    <td><a href="#" onClick="alinearColumna(<%=d%>,2)"><img src="img_tabla/centrar.gif" alt="Centrar celdas" width="22" height="19" border="0"></a></td>
    <td><a href="#" onClick="alinearColumna(<%=d%>,3)"><img src="img_tabla/alinearDer.gif" alt="Alinear celdas a la derecha" width="22" height="19" border="0"></a></td>
  </tr>
</table>

					  
					  <br> 
                      </td>
                      <%
						next
						%>
                    </tr>
                    <%
						%>
                  </table></td>

              </tr>
            </table>
            <%
						set columnas = nothing
						set filas = nothing
					else
						datos = false
					end if
				else
					datos = false
				end if
				if not datos then
					%>
            Insertar tabla de: 
            <input name="filas" type="text" class="campoAdmin" value="3">
            x 
            <input name="columnas" type="text" class="campoAdmin" value="3">
            <input type="button" class="botonAdmin" onClick="nuevatabla(filas.value,columnas.value)" value="Insertar">
            <%
				end if

				%>
<div align="right"> 
  <br>
              <input name="" type="button" class="botonAdmin" onClick="parent.opener.location.href=parent.opener.location;
window.close();" value="Aceptar">
</div>
            <%
		else ' Error, no se encuentra el nodo indicado
			unerror = true
			msgerror = "<br>Error, no se encuentra el nodo indicado"
		end if
	else ' Error, no se encuentra el archivo XML
		unerror = true
		msgerror = "<br>Error, no se encuentra el archivo XML"
	end if
else ' Error, faltan datos.
	unerror = true
	msgerror = "<br>Error, faltan datos. (E1)"
end if

If unerror then
	%>
            <script>
	var texto
	texto = "<%=msgerror%>"
	texto = texto.replace(/<b>|<\/b>|<br>/g,'')
	alert(texto)
	window.history.back()
	</script>
            <%
end if
%>

<%
end select

if request.QueryString("rutavuelta") <> "" then
	rutavuelta = request.QueryString("rutavuelta")
	session("rutavuelta") = rutavuelta
elseif session("rutavuelta") <> "" then
	rutavuelta = session("rutavuelta")
else
	rutavuelta = ""
end if

' Llamada a generahtml (función incluida)
'genera = generahtm(secc,replace(secc,".xml",".asp"),rutavuelta)
%>

		  
		  
		  
		  
		  
</body>
</html>
