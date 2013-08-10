<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include file="usuarios/rutinasParaAdmin.asp" -->
<!--#include file="inc_nav.asp" -->
<%


	secc = replace(""&request.QueryString("secc"),"\","/")
	host = request.ServerVariables("HTTP_HOST")
	url = request.ServerVariables("URL")
	url_actual =  lcase(host & url)

	if inStr(url_actual,host & "/" & c_s & "esp/") then
		idioma = "esp"
	elseif inStr(url_actual,host & "/" & c_s & "eng/") then
		idioma = "eng"
	elseif inStr(url_actual,host & "/" & c_s & "fra/") then
		idioma = "fra"
	elseif inStr(url_actual,host & "/" & c_s & "deu/") then
		idioma = "deu"
	elseif inStr(url_actual,host & "/" & c_s & "ita/") then
		idioma = "ita"
	else
		if session("idioma") <> "" then
			idioma = session("idioma")
		else
			unerror = true : msgerror = "No ha sido posible determinar el idioma de navegación / aSkipper"
		end if
	end if

	if ""&session("cualid") = "edicionmovil" then
		nav_sistema = "movil"
	end if

	if secc="" then
		secc="/inicio"
	end if
	
	ruta_xml_secciones = "/"& c_s & idioma & "/secciones.xml"

	if not unerror then
		session("idioma") = idioma

		xml_nombre = nombreArchivo(secc)
		carpeta = "/" & c_s & idioma & secc
		xml_pagina = carpeta & "/" & xml_nombre & ".xml"

		Set xmlObj = CreateObject("MSXML.DOMDocument")
		if not xmlObj.Load(server.MapPath(xml_pagina)) then
			unerror = true : msgerror = "No se ha podido cargar el XML de secciones*<br>["& xml_pagina &"]."
		else
			set nodoContenido = xmlObj.selectSingleNode("contenido")
			if not typeOK(nodoContenido) then
				unerror = true : msgerror = "El XML de la página actual está incorrecto."
			end if
		end if
	end if


	if not unerror then
		Set xmlSecciones = CreateObject("MSXML.DOMDocument")
		if not xmlSecciones.load( server.MapPath(ruta_xml_secciones) ) then
			unerror = true : msgerror = "No se ha podido cargar el XML de secciones."
		else
			set nodoSeccion = xmlSecciones.selectSingleNode("pagina/secciones" & secc)
			if not typeOK(nodoSeccion) then
				unerror = true : msgerror = "No se ha encontrado el nodo de la sección actual en el XML de secciones." 
			end if
		end if
	end if

	' Zona 1 => navegante (por defecto)
	' Zona 2 => aSkipper
	if not unerror then
		if ""&zona = "" and session("zona") = 2 and ""&nodoSeccion.getAttribute("editable") <> "0" then
			if getPermisoParaRuta("edicion", idioma, session("usuario"),secc) and request.QueryString("nav") <> "movil" then
				zona = 2
			else
				zona = 1
			end if
		else
			zona = 1
		end if
	end if

	if not unerror then
		' Incluimos las funciones para el control de cadenas de texto
		%><script language="javascript" type="text/javascript" src="/<%=c_s%>admin/cadenas.js"></script><%
		if zona = 2 then%>
			<script language="javascript" type="text/javascript">
			<!--
				function winPop(url,winName,ancho,alto,barras) { 
					var winl = (screen.width - ancho) / 2;
					var wint = (screen.height - alto) / 2;
					var paramet = 'top='+wint+',left='+winl+',resizable=yes,width='+ancho+',height='+alto+',scrollbars='+barras+'';
					var sw = window.open(url,winName,paramet);
					sw.focus();
				}
				function volver(){
					window.location.href="/<%=c_s%><%=idioma%>/index.asp?secc=<%=secc%>&salirZona2=1"
				}
				function meta(){
					winPop("/<%=c_s%>admin/edicion/metatags.asp?secc=<%=secc%>&idioma=<%=idioma%>","MetaTags",520,260,0)
				}
				function popMovil(){
					winPop("http://<%=request.ServerVariables("HTTP_HOST")&request.ServerVariables("URL")%>?<%=request.QueryString()%>&nav=movil","DispositivoMovil",176,208,1)
				}
				function propiedades(){
					winPop("/<%=c_s%>admin/edicion/lista.asp?ac=propiedades&nombreactual=<%=nodoSeccion.getAttribute("titulo")%>&estilotitulo=<%=nodoSeccion.getAttribute("estilotitulo")%>&secc=<%=secc%>","Propiedades",403,405,0)
				}
				function plantillas(){
					winPop("/<%=c_s%>admin/edicion/plantilla.asp?secc=<%=secc%>","Plantillas",750,500,1)
				}
				function editar(tipo, num){
					switch (tipo) {
					  case "tabla":
						winPop("/<%=c_s%>admin/edicion/tabla.asp?secc=<%=secc%>&numTabla="+num,"Tabla",700,550,1)
						break;
					  case "imagen":
						winPop("/<%=c_s%>admin/edicion/imagen_frames.asp?secc=<%=secc%>&idi=<%=idioma%>&num="+num,"Imagen",700,550,1)
						//winPop("/<%=c_s%>admin/edicion/imagen.asp?secc=<%=secc%>&idi=<%=idioma%>&num="+num,"Imagen",700,550,1)
						break;
					  case "texto":
						winPop("/<%=c_s%>admin/edicion/texto.asp?archivo=<%=secc%>&idi=<%=idioma%>&num="+num,"Texto",700,550,1)
						break;
					  case "formulario":
					  	// alert("Por el momento, el formulario es de sólo lectura.")
						winPop("/<%=c_s%>admin/edicion/formulario.asp?secc=<%=secc%>&idi=<%=idioma%>&num="+num,"Formulario",450,470,1)
						break;
					  default:
						//
					}		
				}
				function editarCelda(fila,columna,num){
					winPop("/<%=c_s%>admin/edicion/tabla.asp?ac=editar&fila="+ fila +"&columna="+ columna +"&secc=<%=secc%>&numTabla="+ num +"","Tabla",650,450,1)
				}
				function overCeldaAdmin(celda){
					celda.className="celdaAdminOver";
				}
				function outCeldaAdmin(celda){
					celda.className="celdaAdminOut";
				}
				function admingeneral(){
					top.location.href="../admin"
				}
			//-->
			</script>
			<link href="/<%=c_s%>admin/global/estilos.css" rel="stylesheet" type="text/css">
			<link href="/<%=c_s%>estilos.css" rel="stylesheet" type="text/css">
			<style type="text/css">
			<!--
			body {
				margin-left: 0px;
				margin-top: 0px;
				margin-right: 0px;
				margin-bottom: 0px;
			}
		
			.celdaAdminOver {
				background-color: #FFFFE1;
			}
			.celdaAdminOut {
				background-color: #FFFFFF;
			}		

			-->
			</style>
			<table width="100%"  border="0" cellpadding="2" cellspacing="0" bgcolor="#3980F3">
			  <tr>
				<td><span title=" Esta usted en el Modo de aSkipper. Si desa volver al modo normal pulse en el botón Salir ">&nbsp;<b><font color="#FFFFFF">Skipper</font></b> <font color="#FFFFFF">(<%=getNombreIdioma(idioma)%>)</font></span></td>
				<td align="right">
				<%'if ""&nodoSeccion.getAttribute("plantilla") <> "" then%>
					<!--<font color="#FFFFFF" title=""><%'=nodoSeccion.getAttribute("plantilla")%></font>-->
				<%'end if%>
				  <%if insTr(nodoSeccion.getAttribute("compatible"),"movil") then%>
				  <input type="button" class="campoAdmin" title=" Ver esta página compatible con dispositivos móviles " onClick="popMovil()" value="Móvil">
				  <%end if%>
				  <input name="" type="button" class="campoAdmin" title=" Cambiar el título de esta página " onClick="propiedades()" value="Propiedades">
				  <%if ""&nodoSeccion.getAttribute("plantilla")<> "" then%>
				<input name="" type="button" class="campoAdmin" title=" Cambiar la plantilla de esta página " onClick="plantillas()" value="Plantillas">
				<%end if%>
				<input name="" type="button" class="campoAdmin" title=" Ir a la zona de aSkipper general " onClick="admingeneral()" value="Admin">
				</td>
			  </tr>
			</table><br>
		<%end if ' ZONA 2
	end if

	'if secc = "" then
	'	a = Server.MapPath(".")
	'	b = Server.MapPath("/")
	'	ruta = replace(a,b,"")
		
	'	c = 2
	'	ruta = replace(ruta,"\","/")
	'	barras = cuentaPalabras(ruta,"/")-c
	'	for n = 1 to barras
	'		atras = "../" & atras
	'	next
	'	nombre = nombreArchivo(Request.ServerVariables("URL"))
	'end if

	if secc <> "" then
		ficheronuevo = secc &"/"& nombreArchivo(secc)&".xml"
	end if
	
	Function pintaimagen(num)
		dim unerror, msgerror
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if not unerror then
			set nodoImagen = nodoContenido.selectSingleNode("imagen"&num)
			if not typeOK(nodoImagen) then
				unerror = true : msgerror = "El nodo ['contenido/imagen"&num&"'] no se encuentra en el XML de la página actual."
			end if
		end if
		
		if not unerror then
			if ""&nodoImagen.getattribute("novisible") = "1" then
				if not zona = 2 then
					pintaimagen = ""
					exit function
				end if
			end if
			
			Set atributos = nodoImagen.attributes
			if not typeOK(atributos) then
				unerror = true : msgerror = "El nodo imagen no tiene atributos."
			end if
		end if

		if not unerror then
			pie = ""&nodoImagen.getAttribute("pie")
			if pie = "-1" then pie = "" end if
			enlace = ""&nodoImagen.getAttribute("enlace")
			enlaceventana = ""&nodoImagen.getAttribute("enlaceventana")
			abs_imagen = carpeta &"/"& nodoImagen.text
		end if

		if not unerror then
			if ""&nodoImagen.text = "" then
				unerror = true : msgerror = "[No hay imagen]"
			elseif not existe(server.MapPath(abs_imagen)) then
				unerror = true : msgerror = "Imagen no encontrada en el servidor"
			end if
		end if
	
		if not unerror then
			if ""&nodoImagen.getAttribute("novisible") = "1" then%>
				<table width="100%" border="0" cellpadding="1" cellspacing="0" bgcolor="#FF0000">
					<tr><td align="center"><font color="#FFFFFF"><b>No visible</b></font></td></tr>
			</table>
			<%end if
		end if
	
		if not unerror then
			ancho = numero(nodoImagen.getAttribute("ancho"))
			if ancho > 8 then
				ancho_imagen = ancho - 8 ' (-8 por el celpadding de la tabla.)
			end if
			alto = nodoImagen.getAttribute("alto")
			tipo = lcase(""&getExtension(nodoImagen.text))


			
			if nav_sistema="movil" then
				
				if tipo = "swf" or tipo="gif" then
					cadenahtml=""
				else
					
					'PETARÁ EN EL MOMENTO QUE EL NOMBRE TENGA ".GIF" o ".jpg" dentro del nombre
					
					' HAY QUE HACER MAS COMPROBACIONES
					
					archivo=Replace(lcase(abs_imagen),".jpg","_movil.jpg")
					
					if existe(Server.MapPath(archivo)) then
					cadenahtml="<br><a href='"&archivo&"'  border='0'>Ver Imagen</a>"
					end if
				end if
			
			
			
			else
			
			
			if ""&nodoImagen.getattribute("margen") = "1" then
				cadenahtml = "<table border='0' cellpadding='0' cellspacing='0' class='general-foto-margen'><tr><td><table border='0' cellpadding='0' cellspacing='0'><tr><td><img src='../img/foto_s_i.gif'></td><td background='../img/foto_s.gif'><img src='../spacer.gif' width='1' height='1'></td><td><img src='../img/foto_s_d.gif'></td></tr><tr><td background='../img/foto_i.gif'><img src='../spacer.gif' width='1' height='1'></td><td><table cellpadding='0' cellspacing='0' border='0'><tr><td class='general-foto-pixel' align='center'>"
			else
				cadenahtml = "<table cellpadding='0' cellspacing='0' border='0' class='general-foto-margen'><tr><td align='center'>"
			end if

			' Enlace
			if enlace <> "" then
				cadenahtml = cadenahtml & "<a href='"& enlace &"' target='"& enlaceventana &"'>"
			end if

			if tipo = "swf" then
				cadenahtml = cadenahtml & "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width='"& nodoImagen.getattribute("ancho") &"' height='"& nodoImagen.getattribute("alto") &"'><param name='movie' value='"& abs_imagen &"'><param name='quality' value='high'><embed src='"& abs_imagen &"' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='"& nodoImagen.getattribute("ancho") &"' height='"& nodoImagen.getattribute("alto") &"'></embed></object>"
			else
				cadenahtml = cadenahtml & "<img src='"&abs_imagen&"' border='0'>"
			end if

			' FIn de enlace
			if enlace <> "" then
				cadenahtml = cadenahtml & "</a>"
			end if
			cadenahtml = cadenahtml & "</td></tr>"

			if pie <> "" then
				cadenahtml=cadenahtml&"<tr><td align='center'><table width='100%' border='0' cellspacing='0' cellpadding='5' class='general-foto-fondo'><tr><td width='"& ancho_imagen &"' align='center'><b>"& pie &"</b></td></tr></table></td></tr>"
			end if

			if ""&nodoImagen.getattribute("margen") = "1" then
				cadenahtml=cadenahtml&"</table></td><td background='../img/foto_d.gif'><img src='../spacer.gif' width='1' height='1' border='0'></td></tr><tr><td><img src='../img/foto_b_i.gif'></td><td background='../img/foto_b.gif'><img src='../spacer.gif' width='1' height='1'></td><td><img src='../img/foto_b_d.gif'></td></tr></table></td></tr>"
			end if

			cadenahtml = cadenahtml & "</table>"
			

		end if
		if cadenahtml <> "" then
				pintaimagen = cadenahtml
		end if
		end if
		
		if unerror and session("zona") = 2 and msgerror <> "" then
			Response.Write  msgerror
		end if
		
	end Function
	
	sub imagen(num)
		if not unerror then
			if zona = 2 then%>
				<span id="imagen<%=num%>" class="">
				<%=pintaimagen(num)%>
				</span> 
<div align="right"><a href="JavaScript:editar('imagen','<%=num%>')" onMouseOver="overCeldaAdmin(imagen<%=num%>)" onMouseOut="outCeldaAdmin(imagen<%=num%>)"><img src="/<%=c_s%>admin/global/img/lapiz.gif" alt=" Editar imagen " border="0"></a></div>
				<%else
				Response.Write pintaimagen(num)
			end if
		end if
		if unerror then
			if zona = 2 and msgerror <> "" then
				Response.Write msgerror
			end if
		end if
	end sub
	
	Sub texto(num)
	
		if not unerror then
			set nodoTexto = nodoContenido.selectSingleNode("texto"&num)
			if not typeOK(nodoTexto) then
				unerror = true : msgerror = "No se ha encontrado el nodo indicado en el XML<br>['contenido/texto"&num&"']."
			end if
		end if
		
		if not unerror then
			texto_salida = replace(nodoTexto.text,vbCrLf,"<br>")
			' Zona de aSkipper
			if zona = 2 then
				%>
				<span id="texto<%=num%>">
					<%if ""&texto_salida <> "" then%>
						<%=texto_salida%>
					<%else%>
						[Vacio]
					<%end if%>
				</span>
				<div align="right"><a href="JavaScript:editar('texto','<%=num%>')" onMouseOver="overCeldaAdmin(texto<%=num%>)" onMouseOut="outCeldaAdmin(texto<%=num%>)"><img src="/<%=c_s%>admin/global/img/lapiz.gif" alt=" Editar texto " border="0"></a></div>
				<%else
				Response.Write texto_salida
			end if
		end if
		
		if unerror and msgerror <> "" then
			if zona = 2 then
				Response.Write msgerror
			end if
		end if
	
	end Sub
	
	Sub tabla(num)
		if not unerror then
			if zona = 2 then%>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr><td>
					<table width="100%" border="2" cellpadding="0" cellspacing="0" bordercolor="#D5DFE6" id="tabla<%=num%>" style="border:1px">
						<tr><td><%=pintatabla(num)%></td></tr>
					</table></td></tr>
					<tr><td align="right">
					<table border="0" cellpadding="0" cellspacing="0">
						<tr>
						  <td valign="top"><a href="JavaScript:editar('tabla','<%=num%>')" onMouseOver="overCeldaAdmin(tabla<%=num%>)" onMouseOut="outCeldaAdmin(tabla<%=num%>)"><img src="/<%=c_s%>admin/global/img/lapiz.gif" alt=" Editar tabla " border="0"></a>					      </td>
						</tr>
						<tr><td height="2"><img src="../spacer.gif" width="1" height="1"></td></tr>
					</table></td></tr>
				</table>
			<%else
				Response.Write pintatabla(num)
			end if
		end if
		if unerror and msgerror <> "" then
			if zona = 2 then
				Response.Write msgerror
			end if
		end if
	end Sub
	
	function pintatabla(num)
		dim unerror : dim msgerror
	
		set nodoTabla = nodoContenido.selectSingleNode("tabla"& num)
		if not typeOK(nodoTabla) then
			unerror = true : msgerror = "El nodo ['contenido/tabla"&num&"'] no se encuentra en el XML actual."
		end if
		
		if not unerror then
			if nodoTabla.getattribute("visible")=1 or zona = 2 then
				set filas = nodoTabla.childNodes
				color_n = true
				if ""&nodoTabla.getAttribute("visible") = "0" then
					contenido = "<table width=100% border=0 cellpadding=1 cellspacing=0 bgcolor=#FF0000><tr><td align=center><font color=#FFFFFF><b>No visible</b></font></td></tr></table>"
				end if
				contenido = contenido & "<table width='"&nodoTabla.getAttribute("ancho")&"' height='"&nodoTabla.getAttribute("alto")&"' class='general-tabla-ext'><tr><td>"
				contenido = contenido & "<table width='100%' height='100%' class='general-tabla' cellpadding=0 cellspacing=0>"
				for n=0 to filas.length-1
					if color_n then
						color_n = false
						colorCelda = "general-tabla-celdas1"
					else
						color_n = true
						colorCelda = "general-tabla-celdas2"
					end if
					contenido = contenido & "<tr>"
					
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
						
						if zona = 2 then
							botonEditar = "<a href='JavaScript:editarCelda("& n &","& n2 &","& num &")'><img src='/"& c_s &"admin/global/img/lapiz.gif' alt=' Editar celda ' border='0' align='absmiddle'></a>&nbsp;"
						end if

						if columnas.item(n2).getAttribute("resalte") = "1" or  filas.item(n).getAttribute("resalte") = "1" then
							contenido = contenido & "<td class='general-tabla-titulo' align='"&alineado&"'>"&botonEditar&filas.item(n).childNodes.item(n2).text&"</td>"
						else
							contenido = contenido & "<td class='"&colorCelda&"' align='"&alineado&"' valign='top'>"&botonEditar&filas.item(n).childNodes.item(n2).text&"</td>"
						end if
					next
					contenido = contenido & "</tr>"
				next
				contenido = contenido & "</table></td></tr></table>"
				
				set columnas = nothing
				set filas = nothing
			else
				unerror = true : msgerror = "La tabla no es visible."
			end if
		Else ' carga del xml
			unerror = true : msgerror = "Error en nodos."
		end If
	
		if unerror then
			if session("usuario") = 1 then
				pintatabla = msgerror
			else
				pintatabla = ""
			end if
		else
			pintatabla = contenido
		end if
		
	end function

	Sub formulario(num)
		if not unerror then
			
			' Skipper Activo
			if zona = 2 then%>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr>
					<td>
					<table width="100%" border="2" cellpadding="0" cellspacing="0" bordercolor="#D5DFE6" id="formulario<%=num%>" style="border:1px">
						<tr>
						<td><%=getFormulario(num,"")%></td>
						</tr>
					</table>
					</td>
					</tr>
					<tr>
					<td align="right">
					<table border="0" cellpadding="0" cellspacing="0">
						<tr>
						<td valign="top"><a href="JavaScript:editar('formulario','<%=num%>')" onMouseOver="overCeldaAdmin(formulario<%=num%>)" onMouseOut="outCeldaAdmin(formulario<%=num%>)"><img src="/<%=c_s%>admin/global/img/lapiz.gif" alt=" Editar formulario " border="0"></a></td>
						</tr>
						<tr>
						<td height="2"><img src="../spacer.gif" width="1" height="1"></td>
						</tr>
					</table>
					</td>
					</tr>
				</table>
			<%else
				Response.Write getFormulario(num,"")
			end if
		end if

		if unerror and msgerror <> "" then
			if zona = 2 then
				Response.Write msgerror
			end if
		end if
	end Sub
	
	function getFormulario(num, grupo)
		dim unerror, msgerror
		dim salida
		salida = ""

		dim DatosDeUsuario
		DatosDeUsuario = false
		' DatosDeUsuario = true  ==> Los datos de toman del grupo o nodo de un usuario.
		' DatosDeUsuario = false ==> Los datos de toman de la plantilla (fefault).

		' ¿Tomamos los datos de un usuario?
		if typeOK(grupo) then
			set nodoForm = grupo
			DatosDeUsuario = true
		end if

		if not typeOK(nodoForm) then
			set nodoForm = nodoContenido.selectSingleNode("formulario"& num)
			if not typeOK(nodoForm) then
				unerror = true : msgerror = "El nodo ['contenido/formulario"&num&"'] no se encuentra en el XML actual."
			end if
		end if
		
		if not unerror then
			set nodoConfig = nodoForm.selectSingleNode("config")
			if not typeOK(nodoConfig) then
				unerror = true : msgerror = "El nodo ['formulario"&num&"/config'] no se encuentra en el XML actual."
			end if
		end if

		if not unerror then
			if DatosDeUsuario then
				action = "index.asp?secc="& secc
			else
				action = ""&nodoConfig.getAttribute("action")
			end if

			if action="" then
				unerror = true : msgerror = "La configuración general de este formulario no es correcta.<br>Actualice la plantilla."
			end if
		end if
		
		if not unerror and not DatosDeUsuario then

			' Si está activo el atributo "archivar" (por que se lo hemos permitido...)
			' debemos configurar tambien el resto de parametros para lograr archivar
			if ""&nodoForm.getAttribute("archivar")="1" and ""&nodoConfig.getAttribute("archivable")="1" then
				if ""&nodoConfig.getAttribute("cualidad")="" then
					unerror = true : msgerror = "La configuración general de este formulario no es correcta.<br>Actualice la plantilla."
				else
					if ""&nodoForm.getAttribute("seccion")="" then
						unerror = true : msgerror = "Indique la 'Sección' del formulario."
					end if
		
					if ""&nodoForm.getAttribute("idseccion")="" then
						unerror = true : msgerror = "Indique la 'Sección' del formulario."
					end if
				end if
			end if
			
			if ""&nodoForm.getAttribute("asunto")="" then
				unerror = true : msgerror = "Indique el 'Asunto' del formulario."
			end if

			if ""&nodoForm.getAttribute("titulo")="" then
				unerror = true : msgerror = "Indique el 'Título' del formulario."
			end if

		end if
		
		if not unerror then
			if nodoForm.childNOdes.length > 0 then
				salida = leeForm (nodoForm, "", num, action, idioma)
			end if
		end if
		
		if unerror then
			if zona = 2 then
				getFormulario = msgerror
			else
				getFormulario = "" ' Si falla en modo navegación no devuelve nada
			end if
		else
			getFormulario = salida
		end if

	end function
	
	function leeForm (nodoForm, nodoDatos, num, action, idioma)

		Dim js		' Todo el javascript relacionado
		Dim salida	' Etiquetas

		js = ""
		salida = ""

		num = numero(num) ' Identificador de este formulario
		
		' Tomar los datos de un nodo en parametro "nodoDatos" o de un formulario
		dim fuente_en_nodo
		if typeOK(nodoDatos) then
			fuente_en_nodo = true
		else
			fuente_en_nodo = false
		end if

		js = "<script language='javascript' type='text/javascript'>" & vbCrlf
		js = js & "<!--" & vbCrlf
		js = js & "function validar"& num &"(){" & vbCrlf
		js = js & "var s = true" & vbCrlf ' "s" es la variable boolena de salida
		salida = salida & "<form name='f"& num &"' method='post' action='"& action &"' onSubmit='return validar"&num&"()'>" & vbCrlf
		salida = salida & "<input type='hidden' name='form_idioma' value='"& idioma &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='form_seccion' value='"& nodoForm.getAttribute("seccion") &"'>" & vbCrlf
'		salida = salida & "<input type='hidden' name='secc' value='"& secc &"'>" & vbCrlf ' Parece no ser usado
		if ""&nodoForm.getAttribute("FuenteFromName") <> "1" then
			salida = salida & "<input type='hidden' name='FromName' value='"& nodoForm.getAttribute("FromName") &"'>" & vbCrlf
		end if
		if ""&nodoForm.getAttribute("FuenteFromAddress") <> "1" then
			salida = salida & "<input type='hidden' name='FromAddress' value='"& nodoForm.getAttribute("FromAddress") &"'>" & vbCrlf
		end if
		salida = salida & "<input type='hidden' name='idseccion' value='"& nodoForm.getAttribute("idseccion") &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='Recipient' value='"& nodoForm.getAttribute("Recipient") &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='RecipientName' value='"& nodoForm.getAttribute("RecipientName") &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='form_asunto' value='"& nodoForm.getAttribute("asunto") &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='form_titulo' value='"& nodoForm.getAttribute("titulo") &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='msgerror' value='"& nodoForm.getAttribute("msgerror") &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='urlerror' value='"& nodoForm.getAttribute("urlerror") &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='msgexito' value='"& nodoForm.getAttribute("msgexito") &"'>" & vbCrlf
		salida = salida & "<input type='hidden' name='urlexito' value='"& nodoForm.getAttribute("urlexito") &"'>" & vbCrlf
		anchoform = numero(nodoForm.getAttribute("ancho"))
		if anchoform = 0 then
			anchoform = "100%"
		end if
		espaciadobloques = ""&nodoForm.getAttribute("espaciadobloques")
		if espaciadobloques = "" then
			espaciadobloques = 5 ' Espaciado por defecto.
		end if
		salida = salida & "<table border=0 cellpadding=0 cellspacing=0 width='"& anchoform &"' align='"& nodoForm.getAttribute("align") &"'><tr><td>" & vbCrlf
		for each bloque in nodoForm.childNOdes
			if bloque.nodeName = "bloque" then
				salida = salida & "<table border=0 cellpadding=0 cellspacing=0 width='4' height='4'><tr><td><img src='/"& c_s &"spacer.gif' width='4' height=4/></td></tr></table>" & vbCrlf
				if ""&bloque.getAttribute("titulo") <> "" then
					anchoBloque = numero(bloque.getAttribute("ancho"))
					if anchoBloque = 0 then anchoBloque = "100%" end if
					salida = salida & "<table width='"& anchoBloque &"' border=0 cellpadding=0 cellspacing=0><tr><td><fieldset><legend>"& bloque.getAttribute("titulo") &"&nbsp;</legend>" & vbCrlf
				end if
				if bloque.childNOdes.length > 0 then
					espaciado = ""&bloque.getAttribute("espaciado")
					if espaciado = "" then espaciado = 2 end if
					salida = salida & "<table align='"& bloque.getAttribute("align") &"' width='100%' border='0' cellspacing='0' cellpadding="& espaciadobloques &"><tr><td>" & vbCrlf
					salida = salida & "<table width='100%' border='0' cellspacing='0' cellpadding='"& espaciado &"'>" & vbCrlf
				for each fila in bloque.childNOdes	
						' FILA
						if fila.nodeName = "fila" then
							if ""&fila.getAttribute("titulo")<>"" then
								salida = salida & "<tr><td align='"& fila.getAttribute("align") &"'><nobr>" & fila.getAttribute("titulo") & "</td></tr>" & vbCrlf
							end if
							salida = salida & "<tr><td align='"& fila.getAttribute("align") &"'><nobr>" & vbCrlf
							for each campo in fila.childNodes
								if ""&campo.getAttribute("requerido") = "1" then
									select case ""&campo.getAttribute("validar")
									case "textovacio",""
										js = js & "if (s) { s = noVacio('"&campo.getAttribute("nombrecorto")&"','"& campo.getAttribute("titulo") &"') }" & vbCrlf
									case "numero"
										js = js & "if (s) { s = esNumero('"&campo.getAttribute("nombrecorto")&"','"& campo.getAttribute("titulo") &"') }" & vbCrlf
									case "email"
										js = js & "if (s) { s = esEmail('"&campo.getAttribute("nombrecorto")&"','"& campo.getAttribute("titulo") &"') }" & vbCrlf
									end select
								end if

								' Tomo el valor del nodo xml o de un formularo (es posible que de ninguno)
								if fuente_en_nodo then
									' Defino el valor que hay en el nodo
									set nodo_campo_dato = nodoDatos.selectSingleNode(campo.getAttribute("nombrecorto"))
									if typeOK(nodo_campo_dato) then
										valor = ""& nodo_campo_dato.text
									end if
								else
									valor = request.Form(campo.getAttribute("nombrecorto"))
								end if
								
								salida = salida & campoForm (campo,valor) & "" & vbCrlf
							next
							salida = salida & "</nobr></td></tr>" & vbCrlf
						end if
						
						' SALTO DE LINEA
						if fila.nodename = "saltodelinea" then
							if ""&fila.getAttribute("lineahorizontal") = "1" then
								salida = salida & "<tr><td><hr width=100% size=1 noshade></td></tr>" & vbCrlf
							else
								salida = salida & "<tr><td>&nbsp;</td></tr>" & vbCrlf
							end if
						end if
					next
					salida = salida & "</table>" & vbCrlf
					salida = salida & "</td></tr></table>" & vbCrlf ' tabla para aplicar cellpaddin
				end if
				if ""&bloque.getAttribute("titulo")<>"" then
					salida = salida & "</fieldset></td></tr></table>" & vbCrlf
				end if
			end if ' nodename: bloque
		next
		salida = salida & "<table border=0 cellpadding=0 cellspacing=0 width='4' height='4'><tr><td><img src='/"& c_s &"spacer.gif' width='4' height=4/></td></tr></table>" & vbCrlf
		salida = salida & "<table width='"& anchoform &"' border=0 cellpadding=2 cellspacing=0><tr><td align='right'><input name='' type='image' src='/"& c_s &  idioma &"/imagenes/boton_enviar.gif' border='0'></td></tr></table>" & vbCrlf
		salida = salida & "</td></tr></table>" & vbCrlf ' tabla primera
		salida = salida & "</form>" & vbCrlf

		js = js & "return s" & vbCrlf
		js = js & "}" & vbCrlf
		js = js & "//-->" & vbCrlf
		js = js & "</script>" & vbCrlf

		leeForm = salida & js

	end function
	
	function getTexto(num)
		if not unerror then
			set nodoTexto = nodoContenido.selectSingleNode("texto"&num)
			if typeOK(nodoTexto) then
				getTexto = ""&nodoTexto.text
			end if
		end if
	end function

	function campoForm(campo,valor)
		dim s : s = ""
		set atts = campo.attributes
		for each att in atts
			sAtt = sAtt &" "& att.nodeName &"='"& att.text &"'"
		next
		nombrecorto = ""&campo.getAttribute("nombrecorto")
		if campo.parentNode.parentNode.parentNode.getAttribute("FuenteFromName") = "1" and nombrecorto = ""&campo.parentNode.parentNode.parentNode.getAttribute("FromName") then
			nombrecorto = "FromName"
		end if
		if campo.parentNode.parentNode.parentNode.getAttribute("FuenteFromAddress") = "1" and nombrecorto = ""&campo.parentNode.parentNode.parentNode.getAttribute("FromAddress") then
			nombrecorto = "FromAddress"
		end if
		select case campo.getAttribute("tipo")
			case "oculto"
				s = s &"<input name='"& nombrecorto &"' id='"& nombrecorto &"' type='hidden'>"
			case "texto"
				if ""&campo.getAttribute("titulo")<>"" then
					s = s &"<b>"& campo.getAttribute("titulo") &"</b>"
					if ""&campo.getAttribute("requerido") = "1" then
						s = s &"*: "
					else
						s = s &": "
					end if
				end if
				s = s &"<input name='"& nombrecorto &"' id='"& nombrecorto &"' type='text' class='campo'"& sAtt &" value='"& valor &"'>" & vbCrlf
			case "clave"
				if ""&campo.getAttribute("titulo")<>"" then
					s = s &"<b>"& campo.getAttribute("titulo") &"</b>"
					if ""&campo.getAttribute("requerido") = "1" then
						s = s &"*: "
					else
						s = s &": "
					end if
				end if
				s = s &"<input name='"& nombrecorto &"' id='"& nombrecorto &"' type='password' class='campo'"& sAtt &" value='"& valor &"'>" & vbCrlf
			case "memo"
				if ""&campo.getAttribute("titulo") <> "" then
					s = s &"<nobr><b>"& campo.getAttribute("titulo") &"</b>:</nobr><br>" & vbCrlf
					if ""&campo.getAttribute("requerido") = "1" then
						s = s &"*: "
					else
						s = s &": "
					end if
				end if
				s = s &"<textarea name='"& nombrecorto &"' class='area'"& sAtt &"></textarea>" & vbCrlf
			case "combo" ' Desplegable
				s = s &"<b>"& campo.getAttribute("titulo") &"</b>"
				if ""&campo.getAttribute("requerido") = "1" then
					s = s &"*: "
				else
					s = s &": "
				end if
				s = s & "<select name='"& nombrecorto& "' class='campo'"& sAtt &">"
				for each opcion in campo.childnodes
					if (valor = "" and ""&opcion.getAttribute("default") = "1") or valor = opcion.getAttribute("valor") then
						default = " selected"
					else
						default = ""
					end if
					s = s & "<option value='"& opcion.getAttribute("valor") &"'"& default &">"& opcion.getAttribute("titulo") &"</option>"
				next
				s = s & "</select>"
			case "opcion" ' Radio button
				s = s &"<nobr><b>"& campo.getAttribute("titulo") &"</b>:</nobr> "
				for each opcion in campo.childnodes
					if (valor = "" and ""&opcion.getAttribute("default") = "1") or valor = opcion.getAttribute("valor") then
						default = " checked"
					else
						default = ""
					end if
					s = s & "<input "& sAtt &" name='"& nombrecorto & "' id='"& campo.getAttribute("titulo")&opcion.getAttribute("valor")& "' value='"& opcion.getAttribute("valor") &"' type='radio'"& default &"><label for='"& campo.getAttribute("titulo")&opcion.getAttribute("valor") &"'>"& opcion.getAttribute("titulo") &"</label> "
				next
			case "check"
				s = s &"<nobr><b>"& campo.getAttribute("titulo") &"</b></nobr>: "
				n_opt = 0
				valor = ","& replace(valor,", ",",") &","
				if valor = ",," then
					valor = ""
				end if
				for each opcion in campo.childnodes
					n_opt = n_opt + 1
					if (valor = "" and ""&opcion.getAttribute("default") = "1") or inStr(valor,","& trim(opcion.getAttribute("valor")) &",") then
						default = " checked"
					else
						default = ""
					end if
					s = s & "<input "& sAtt &" name='"& nombrecorto & "' id='"& campo.getAttribute("titulo")&n_opt& "' value='"& opcion.getAttribute("valor") &"' type='checkbox'"& default &"><label for='"& campo.getAttribute("titulo")&n_opt &"'>"& opcion.getAttribute("titulo") &"</label> "
				next
		end select
		campoForm = s
	end function
%>