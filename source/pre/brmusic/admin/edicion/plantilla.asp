<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/inc_rutinas.asp" -->
<html>
<head>
<title>Plantillas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo9 {
	color: #009933;
	font-weight: bold;
}
-->
</style>
</head>
<body class="bodyAdmin">
<%
function getValorD(pV,tipo)
	t = numero(eval(tipo &"_"& pV &"_cuerpo_sitio"))
	if t>0 then
		getValorD = t
	else
		getValorD = numero(pV)
	end if
end function

if session("idioma") = "" then
	unerror = true : msgerror = "Se ha perdido la sesión."
end if

if not unerror then

' Secc
secc = ""&request("secc")
if secc = "" then
	unerror = true : msgerror = "[secc] está vacio."
end if

carpeta = "/" & c_s & session("idioma") & secc

if not unerror then
	set xmlSecciones = CreateObject("MSXML.DOMDocument")
	if not xmlSecciones.load(server.MapPath("/" & c_s & session("idioma") & "/secciones.xml")) then
		unerror = true : msgerror = "No se ha podido cargar el XML secciones."
	end if
end if

if not unerror then
	set nodoSeccion = xmlSecciones.selectSingleNode("/pagina/secciones"& secc)
	if typeName(nodoseccion) = "Nothing" then
		unerror = true : msgerror = "El nodo [pagina/secciones"& secc &"] no se encuentra en el XML secciones."
	end if
end if

select case request("ac")
case "cambiar"
	
	actual = ""&request.QueryString("actual")
	nueva = ""&request.QueryString("nueva")
	secc = ""&request.QueryString("secc")
	informe = ""
	if actual = "" or nueva = "" or secc = "" then
		unerror = true : msgerror = "No se han recibido todos los parametros necesarios."
	end if
	if ""&request.QueryString("c") = "1" then confirma = true else confirma = false end if

	carpeta = "/" & c_s & session("idioma") & secc

	' Plantilla actual
	if not unerror then
		set xmlPlantillaActual = CreateObject("MSXML.DOMDocument")
		if not xmlPlantillaActual.load(server.MapPath(carpeta & "/" & nombreArchivo(secc) & ".xml")) then
			unerror = true : msgerror = "No se ha podido cargar la plantilla actual."
		else
			set plantillaActual = xmlPlantillaActual.selectSingleNode("contenido")
			if not typeOK(plantillaActual) then
				unerror = true : msgerror = "No se ha podido cargar el nodo [contenido] de la plantilla actual."
			end if
		end if
	end if

	' Plantilla nueva
	if not unerror then
		set xmlPlantillaNueva = CreateObject("MSXML.DOMDocument")
		if not xmlPlantillaNueva.load(server.MapPath("/" & c_s & "plantillas/"& nueva & ".xml")) then
			unerror = true : msgerror = "No se ha podido cargar la plantilla nueva.<br>"&server.MapPath("/" & c_s & "plantillas/"& nueva & ".xml")
		else
			set plantillaNueva = xmlPlantillaNueva.selectSingleNode("contenido")
			if not typeOK(plantillaNueva) then
				unerror = true : msgerror = "No se ha podido cargar el nodo [contenido] de la plantilla nueva."
			end if
		end if
	end if

	' Plantilla vacia
	if not unerror then
		set xmlPlantillaVacia = CreateObject("MSXML.DOMDocument")
		if not xmlPlantillaVacia.load(server.MapPath("/" & c_s & "plantillas/vacio.xml")) then
			unerror = true : msgerror = "No se ha podido cargar la plantilla vacia."
		else
			set plantillaVacia = xmlPlantillaVacia.selectSingleNode("contenido")
			if not typeOK(plantillaVacia) then
				unerror = true : msgerror = "No se ha podido cargar el nodo [contenido] de la plantilla vacia."
			end if
		end if
	end if

	if not unerror then
		' Busco los nodos de la nueva en la actual
		for each nodoPNueva in plantillaNueva.childNodes
			set nodoPActual = plantillaActual.selectSingleNode(nodoPNueva.nodeName)
'			if confirma then
				' Tomo el nodo de la Actual si existe, si no, de la nueva (en vacio)
				if typeOK(nodoPActual) then
					' NODO IMAGEN ***
					if inStr(nodoPActual.nodename,"imagen") then
						' Compruebo los atributos
						ancho = numero(nodoPActual.getAttribute("ancho"))
						alto = numero(nodoPActual.getAttribute("alto"))
						anchoMax_actual = getValorD(nodoPActual.getAttribute("anchomax"),"ancho")
						anchoMax_nuevo = getValorD(nodoPNueva.getAttribute("anchomax"),"ancho")
						' Tipo de contenido (SWF, JPG, GIF, ...)
						tipo = Ucase(getExtension(""&nodoPActual.text))
'						informe = informe & "<br>"
'						informe = informe & "<br><b>Tipo:</b> "& tipo 
'						informe = informe & "<br><b>Ancho Actual:</b> "& ancho
'						informe = informe & "<br><b>Alto Actual:</b> "& alt
'						informe = informe & "<br>"
						if ancho > anchoMax_nuevo then
							if tipo = "SWF" then
								informe = informe & "<br><font color='#BB0000'>La <b>" & nodoPActual.nodeName & "</b> es un FLash. No cambiará de tamaño. Revise posibles fallos.</font>"
							else
								informe = informe & "<br><font color='#BB0000'>La <b>" & nodoPActual.nodeName & "</b> se ajustará al nuevo máximo: "& anchoMax_nuevo &" px.</font>"
							end if
							if confirma and tipo <> "SWF" then
								' Aplicar el nuevo tamaño a la foto (ASPJpeg)
								Set jpeg = Server.CreateObject("Persits.Jpeg")
								jpeg.Open(server.MapPath(carpeta & "/" & nodoPActual.text))
								ancho = anchoMax_nuevo
								alto = round((alto*anchoMax_nuevo)/ancho)
								if ancho > 0 and alto > 0 then
									jpeg.Width = ancho
									jpeg.Height = alto
								end if
								jpeg.Save server.MapPath(carpeta & "/" & nodoPActual.text)
								set jpeg = nothing
							end if
						end if
						if confirma then
							' Asignar nuevos atributos al nodo
							set att = xmlPlantillaActual.createAttribute("ancho")
							nodoPActual.setAttributeNode(att)
							att.nodeValue = ancho
							set att = nothing
							
							set att = xmlPlantillaActual.createAttribute("alto")
							nodoPActual.setAttributeNode(att)
							att.nodeValue = alto
							set att = nothing

							set att = xmlPlantillaActual.createAttribute("anchomax")
							nodoPActual.setAttributeNode(att)
							att.nodeValue = anchoMax_nuevo
							set att = nothing
						end if
						plantillaVacia.appendChild(nodoPActual)
					end if

					' NODO FORMULARIO ***
					if inStr(lcase(nodoPActual.nodename),"formulario")>0 then
						informe = informe &"<b><b>He encontrado un formulario!</b></b>"
						plantillaVacia.appendChild(nodoPNueva)
					end if

					' NODO TEXTO ***
					if inStr(nodoPActual.nodename,"texto") then
						plantillaVacia.appendChild(nodoPActual)
					end if

					' NODO TABLA ***
					if inStr(nodoPActual.nodename,"tabla") then
						plantillaVacia.appendChild(nodoPActual)
					end if

					'plantillaVacia.appendChild(nodoPActual)
				else
					plantillaVacia.appendChild(nodoPNueva)
				end if
'			end if
		next

		for each nodoPActual in plantillaActual.childNodes
			set nodoPNueva = plantillaNueva.selectSingleNode(nodoPActual.nodeName)
			if not typeOK(nodoPNueva) then
	
				if inStr(nodoPActual.nodeName,"texto") then
					informe = informe & "<br><font color='#BB0000'>El <b>" & nodoPActual.nodeName & "</b> se perderá.</font>"
				elseif inStr(nodoPActual.nodeName,"imagen") then
					informe = informe & "<br><font color='#BB0000'>La <b>" & nodoPActual.nodeName & "</b> se perderá.</font>"
					if confirma then
						set fso = Server.CreateObject("Scripting.FileSystemObject")
						if existe(server.MapPath(carpeta & "/" & nodoPActual.text)) then
							fso.DeleteFile server.MapPath(carpeta & "/" & nodoPActual.text)
						end if
						set fso = nothing
					end if
				elseif inStr(nodoPActual.nodeName,"tabla") then
					informe = informe & "<br><font color='#BB0000'>La <b>" & nodoPActual.nodeName & "</b> se perderá.</font>"
				elseif inStr(nodoPActual.nodeName,"formulario") then
					informe = informe & "<br><font color='#BB0000'>Se perderá los datos de <b>" & nodoPActual.nodeName & "</b>.</font>"
				end if
	
			end if
		next
		if confirma then
			xmlPlantillaVacia.save Server.MapPath("/" & c_s & session("idioma") & secc & "/" & nombreArchivo(secc) & ".xml")
			set att = xmlSecciones.createAttribute("plantilla")
			nodoSeccion.setAttributeNode(att)
			att.nodeValue = nueva
			xmlSecciones.save server.MapPath("/" & c_s & session("idioma") & "/secciones.xml")
			%>
			Cambio realizado con exito.<br>
			<script language="javascript" type="text/javascript">
				if (parent.opener.name == "principal") {
					parent.opener.location.href = "/<%=c_s%>plantillas/<%=nueva%>.asp?zona=2&secc=<%=secc%>"
				} else {
					parent.opener.location.href = parent.opener.location
				}
				window.close()
			</script>
			<%
		else
		%>
			<script language="javascript" type="text/javascript">
				function cancelarCambio() {
					location.href="plantilla.asp?secc=<%=secc%>"
				}
				function confirmarCambio() {
					location.href="plantilla.asp?ac=cambiar&c=1&secc=<%=secc%>&nueva=<%=nueva%>&actual=<%=actual%>"
				}
			</script>
				<span class="tituloazonaAdmin">Informe y confirmaci&oacute;n</span><br>
			Lea con atenci&oacute;n el siguiente informe para conocer los cambios<br>
			que se producir&aacute;n al realizar el cambio de plantilla.<br>
			<br> 
			De lo contrario, es posible que se borren im&aacute;genes que usted no desea borrar. <br>
			<br>
			<table  border="0" align="center" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
			  <tr>
				<td><font size="+1"><%=nodoSeccion.getAttribute("titulo")%></font>
				<table width="100%"  border="0" align="center" cellpadding="4" cellspacing="0">
				  <tr>
					<td align="center"><span class="tituloazonaAdmin">Actual</span><br><font color="#999999"><%=actual%></font></td>
					<td align="center"><span class="tituloazonaAdmin">Nueva</span><br>
				    <font color="#999999"><%=nueva%></font></td>
				  </tr>
				  <tr>
					<td align="center"><img src="/<%=c_s%>plantillas/<%=actual%>_n.gif"></td>
					<td align="center"><img src="/<%=c_s%>plantillas/<%=nueva%>_n.gif"></td>
				  </tr>
				</table>
				  <table border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#ECE9D8">
					<tr>
					  <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
						  <tr>
							<td width="15" bgcolor="#FFFFFF">&nbsp;</td>
							<td bgcolor="#FFFFFF"><br>
								<span class="Estilo5">INFORME</span></td>
							<td width="15" bgcolor="#FFFFFF">&nbsp;</td>
						  </tr>
						  <tr>
							<td bgcolor="#FFFFFF">&nbsp;</td>
							<td bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="2" cellspacing="0">
								<tr>
								  <td><%if informe <> "" then
								  Response.Write informe
								  else%>
								    <span class="Estilo9">Las plantillas con compatibles</span>								  <%end if%></td>
								</tr>
							  </table>
								<br></td>
							<td bgcolor="#FFFFFF">&nbsp;</td>
						  </tr>
					  </table></td>
					</tr>
				  </table>
				  <table width="100%"  border="0" align="center" cellpadding="2" cellspacing="0">
					<tr>
					  <td align="center"><br>
						<input name="" type="button" class="botonAdmin" onClick="confirmarCambio()" value="Realizar cambio">
						  <input name="" type="button" class="botonAdmin" onClick="cancelarCambio()" value="Cancelar"></td>
					</tr>
			    </table></td>
			  </tr>
			</table>
			<br>
<%		end if
	end if
case else

if not unerror then

	tituloPagina = nodoseccion.getAttribute("titulo")
	plantillaActual = ""&nodoseccion.getAttribute("plantilla")
%>
	<script language="javascript" type="text/javascript">
		function eleccion(plant) {
			location.href='plantilla.asp?ac=cambiar&nueva='+plant+"&secc=<%=secc%>&actual=<%=plantillaActual%>"
		}
	</script>

	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="0">
		<tr> 
		<td><span class="tituloazonaAdmin">Plantillas</span><br>
		La p&aacute;gina actual (<%=tituloPagina%>) tiene la plantilla <%=plantillaActual%>.</td>
	</table>
	
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="6" bgcolor="#FFFFFF">
		<tr>
<%carpetaplant = Server.MapPath("/" & c_s & "plantillas")
inicial=1
set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set Upload = Server.CreateObject("Persits.Upload")
if FSO.FolderExists(carpetaplant) then
	Set Dir = Upload.Directory( carpetaplant&"/*.*", SORTBY_TYPE)
	n=1
	For Each item in Dir
		bFile = carpetaplant&"/"&item.FileName
		if Strcomp(right(item.FileName,4),".asp")=0 then
			ficheroenxml = carpetaplant&"/"&replace(item.FileName,".asp",".xml")
			if Upload.FileExists (ficheroenxml) then
				nombreplant = replace(item.FileName,".asp","")
				if lcase(plantillaActual) <> lcase(nombreplant) then
					actual = true
				else
					actual = false
				end if

				%>
				<td align="center">

	<%if actual then%>
		<table border='0' cellspacing='0' cellpadding='3'>
			<tr> 
			<td align="center"><a href="javascript:eleccion('<%=nombreplant%>')"><img src="/<%=c_s%>plantillas/<%=nombreplant%>.gif" border="0"></a>
			<%=nombreplant%></td>
			</tr>
		</table>
	<%else%>
		<table border='2' cellpadding='3' cellspacing='0' bordercolor='#0066CC'>
			<tr> 
			<td align="center"><a href="javascript:eleccion('<%=nombreplant%>')"><img src="/<%=c_s%>plantillas/<%=nombreplant%>.gif" border="0"></a>
			<%=nombreplant%></td>
			</tr>
		</table>
	<%end if%>
				

				</td>
				<%n=n+1
				if n mod 7 = 1 then%>
	  </tr>
					<tr>
				<%end if
			end if
		end if
	Next
	
	
	modulo=n mod 5
	if modulo <> 1 then
	for p=modulo to 6-modulo
		%>
          <td><nobr> </nobr></td>
          <%
	next
	end if
end if
Set Upload = Nothing
Set FSO = Nothing%>
        </tr>
</table>
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="6">
        <tr> 
          <td align="right" valign="bottom"> <%if session("novolver")<>""then%> <br> <input type="button" class="botonAdmin" onClick="window.close()" value="Cancelar"> 
            <%end if%> </td>
        <tr> 
</table>
<%end if

end select

end if

if unerror then
	Response.Write "<b>Error:</b> <br>"& msgerror
end if
%>
</body>
</html>