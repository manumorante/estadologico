<%
	' Idioma
	pi = request.ServerVariables("PATH_INFO")
	if inStr(pi,"/esp/") then
		idioma = "esp"
	elseif inStr(pi,"/eng/") then
		idioma = "eng"
	elseif inStr(pi,"/fra/") then
		idioma = "fra"
	elseif inStr(pi,"/deu/") then
		idioma = "deu"
	elseif inStr(pi,"/ita/") then
		idioma = "ita"
	else
		if session("idioma") <> "" then
			idioma = session("idioma")
		end if
	end if
%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include file="inc_rutinas.asp" -->
<!--#include virtual="/esp/propios/tituloEstilos/tituloEstilos.asp" -->
<!--#include file="global/inc_lanza_popup.asp" -->
<%

	secc = replace(""&request.QueryString("secc"),"\","/")

	while inStr(secc,"//") 
		secc = replace(secc,"//","/")
	wend
	raiz = "/pagina/secciones"

	if secc = "" then
		if ""&nav_sistema = "movil" then
			response.Redirect("index.asp?secc=/iniciomovil")
		else
			response.Redirect("index.asp?secc=/inicio")
		end if
	end if
	

	' Carga del XML de SECCIONES.
	Dim xmlObj
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if not xmlObj.Load(server.MapPath("/" & c_s & idioma & "/secciones.xml")) then
		unerror = true : msgerror = "No se ha encontrado el XML secciones o está corrupto."
	End If

	if not unerror then
		on error resume next
		' Obtenemos la sección actual.
		dim seccionactual
		set seccionactual = xmlObj.selectSingleNode(raiz & secc)
		if not typeOK(seccionactual) or err<>0 then
			if session("usuario") = 1 then
				unerror = true : msgerror = "La sección indicada no se encuentra."
			else
				Response.Redirect("/" & c_s & idioma & carpeta_adicional & "/index.asp?secc=/mensajes/noexiste")
			end if
		end if
		on error goto 0
	end if

	if not unerror then
		tienetitulo = ""&seccionactual.getAttribute("estilotitulo")	
		tienebotonera = ""&seccionactual.getAttribute("botonera")
		if ""&tienebotonera="" and ""&tienebotonera<>"0" then   'Si el no tiene botonera pero la tiene el padre.. si vale.
		tienebotonera = ""&seccionactual.parentNode.getAttribute("botonera")
		end if
		' Si el atributo "default" es distinto de 0 nos redirigirnos al hijo en la posición indicada.
		default = seccionactual.getAttribute("default")
		if default <> 0 and ""&default <> "0" and default <> "" then
			Dim elNodeName
			elNodeName = xmlObj.selectSingleNode(raiz & secc).childnodes.item(default-1).nodename
			set seccionactual = nothing
			set xmlObj = nothing
			response.redirect("index.asp?secc="&secc&"/"&elNodeName)
		end if
		
		' Composicion del titulo que aparece en la ventana del navegador. (<title></title>)
		
		' CABECERA ___ (en html) (dentro del html le podemos poner un flash)
		' Si la sección en la que estamos no tiene definido una cabecera tomaremos la de su padre
		' si no queremos que coja la de su padre ponemos [urltitulo="0"]
		Dim urltitulo
		Dim urlTituloPadre
		set yo = xmlObj.selectSingleNode(raiz & secc)
		if typeOK(yo) then
			urltitulo = ""&yo.getAttribute("urltitulo")
			urlTituloPadre = ""&yo.parentNode.getAttribute("urltitulo")
		end if
		if urltitulo <> "" then
			if urltitulo <> "0" then
				urltitulo = seccLimpia &"/" & urltitulo
				else
				urltitulo = ""
			end if
		elseif urlTituloPadre <> "" then
			urltitulo = seccLimpia &"/../" & urlTituloPadre
			else
			urltitulo = ""
		end if
	
		' Estilo Titulo	- tomar el atributo del nodo
		dim estilotitulo
		estilotitulo = ""&yo.getAttribute("estilotitulo")	
	end if
	
	if not unerror then
		RutaXmlPagina = "/" & c_s & idioma & carpeta_adicional & secc & "/" & nombreArchivo(secc) & ".xml"
		Set xmlPagina = CreateObject("MSXML.DOMDocument")
		if xmlPagina.Load(server.MapPath(RutaXmlPagina)) then
			Set nodoContenido = xmlPagina.selectSingleNode("contenido")
			if typeOK(nodoContenido) then
				elTitulohtml = titulohtml(seccionactual)
				set nodoTitulo = nodoContenido.selectsinglenode("titulo")
				if typeOK(nodoTitulo) then
					if nodoTitulo.text <> "" then
						titulo = titulo &" - "& nodoTitulo.text
					elseif elTitulohtml <> "" then
						titulo = titulo &" - "& elTitulohtml
					end if
				elseif elTitulohtml <> "" then
					titulo = titulo &" - "& elTitulohtml
				end if
				
				set nodoDescripcion = nodoContenido.selectsinglenode("descripcion")
				if typeOK(nodoDescripcion) then
					descripcion = ""&nodoDescripcion.text
				end if

				set nodoKeys = nodoContenido.selectsinglenode("keys")
				if typeOK(nodoKeys) then
					keys = nodoKeys.text
				end if
			end if
		end if
	end if

	' Funciones -----------------------------------------------------------------------------------

	' Inclusión de elementos externos/especiales como: noticias, cabereras, pie de pagina ....:
	' Funcion "incluirEspecial": recibe el número que representa su posicion,
	' si existe algun atributo con el pues incluye el valos del atributo.
	
	function incluirElemento(nombre_att)
		if typeOK(yo) then
			Dim inc_archivo
			inc_archivo = ""&yo.getAttribute(nombre_att)
	
			if inc_archivo <> "" then
				if existe(server.MapPath(inc_archivo)) then
					server.execute(inc_archivo)
					incluirElemento=true
				else
					incluirElemento=false
				end if
			else
				incluirElemento=false
			end if
		end if
	end function

	' Portada editable
	if ""&cabera_editable = "1" then
		secc_portada = "/"& c_s & "esp/portada"
		ruta_xml_portada = secc_portada &"/portada.xml"
		set xml_portada = CreateObject("MSXML.DOMDocument")
		if not xml_portada.Load(Server.MapPath(ruta_xml_portada)) then
			unerror_temp = true : msgerror_tem = "No se encontrado el XML de portada."
		else
			set nodoImagen = xml_portada.selectSingleNode("contenido/imagen1")
			if not typeOK(nodoImagen) then
				portada_editable = secc_portada&"/por_defecto.jpg"
			else
				cabecera_editable_nombre = ""& nodoImagen.text
				if cabecera_editable_nombre = "" then
					portada_editable = secc_portada&"/por_defecto.jpg"
					cabecera_editable_tipo = "jpg"
				else
					cabecera_editable_enlace = ""&nodoImagen.getAttribute("enlace")
					cabecera_editable_ventana = ""&nodoImagen.getAttribute("ventana")
					cabecera_editable_tipo = lcase(""&getExtension(cabecera_editable_nombre))
					ancho_editable=""&nodoImagen.getAttribute("ancho")
					alto_editable=""&nodoImagen.getAttribute("alto")
					portada_editable = secc_portada&"/"&cabecera_editable_nombre
				end if

			end if
		end if
		
		set xml_portada = nothing
	end if

	Sub Cuerpo()
		
		if ""&seccionactual.getattribute("activo")="0" then
			response.Redirect("index.asp?secc=/mensajes/paginadesactivada")
		end if
		
		' Vemos si tiene plantilla
		Dim plantilla
		plantilla = ""&seccionactual.getattribute("plantilla")
		if plantilla <> "" then
			pagPlantilla = "/"& c_s &"plantillas/"& plantilla &".asp"
			if existe(server.MapPath(pagPlantilla)) then
				server.Execute(pagPlantilla)
			else
				if session("usuario") = 1 then
					Response.Write("InfoAdmin:<br>No se ha encontrado la plantilla indicada.")
				else
					Response.Redirect("index.asp?secc=/mensajes/noexiste&msgerror=No se ha encontrado la plantilla indicada.")
				end if
			end if
		else
			' DESVIO ___
			' si existe un desvio, en modo navegación, ejecuto la pagina indicada en url
			' (la ruta en desvio será de tipo nodo y se usará para la edición)
			
			desvio = ""&xmlObj.selectSingleNode(raiz & secc).getAttribute("desvio")
			url = ""&xmlObj.selectSingleNode(raiz & secc).getAttribute("url")
			if desvio <> "" and url <> "" then
				' Construimos la ruta hasta el archivo indicado en el atributo "url"
				ladireccion = "/" & c_s & idioma & secc &"/"& url
				if existe(server.MapPath(ladireccion)) then 
					server.Execute(ladireccion)
				else
					if session("usuario") = 1 then
						Response.Write("InfoAdmin:<br>La página indicada ("& ladireccion &") no existe.")
					else
						Response.Redirect("index.asp?secc=/mensajes/noexiste")
					end if
				end if
			else
				' NOMBRE NODO ___
				' busco la pagina .asp con el nombre del nodo
				ladireccion = "/" & c_s & idioma & carpeta_adicional & secc & "/" & nombreArchivo(secc) & ".asp"
				if existe(server.MapPath(ladireccion)) then 
					server.Execute(ladireccion)
				else
					if session("usuario") = 1 then
						Response.Write("InfoAdmin:<br>La página indicada ("& ladireccion &") no existe.")
					else
						Response.Redirect("index.asp?secc=/mensajes/noexiste")
					end if
				end if
			end if
		end if
	
	End Sub
	
	' Pita el código que llama a un flash...DOCUMENTAR
	' Hay que modificar la versión según la que sea y dependiendo de la página que lo use.
	
	Sub flash(pelicula,parametros,width,height,color)
		if nav_sistema<>"movil" then ' desactiva los flash en los móviles
		
		if color = "transparente" then
			transparente = true
			color = ""
		end if
		%>
		<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="<%=width%>" height="<%=height%>">
		<param name="movie" value="<%=pelicula%>?<%=parametros%>">
		<param name="quality" value="high">
		<embed src="<%=pelicula%>?<%=parametros%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="<%=width%>" height="<%=height%>"></embed>
		<%if transparente then%>
			<param name="wmode" value="transparent">
		<%end if%>
		</object>
		
		
		<%
		end if
	end sub
	
	' rutamapa
	Sub rutamapa
		dim n, letra, variable, total
		dim esta_seccion
		for n=1 to len(secc) 
			letra = Mid(secc,n,1)
			if letra = "/" then
				if variable = "" then
				else
					total = total&"/"&variable
					variable=""
					if ""&xmlObj.selectsinglenode(raiz & total).getattribute("activo")="0" then
						pintasecsin(total)
					else
						pintasec(total) 	
					end if
				end if
			else
				variable = variable & letra
			end if
		next
		
		if variable<>"" then
			pintasecsin( total & "/" & variable)
		end if
		response.Write("&nbsp;")
	end sub
	
	Sub pintasec(parametro)
		titulo=xmlObj.selectsinglenode(raiz & parametro).getattribute("titulo")
		response.Write("<img src=../img/flecha.gif><font size=1><a href='index.asp?secc="&parametro&"' target='_top'>" & titulo & "</a></font>")
	end Sub
		
	Sub pintasecsin(parametro)	
		dim titulo
		if ""&parametro<>"" then
			set titulo = xmlObj.selectsinglenode(raiz & parametro)
			if typeOK(titulo) then
				eltitulo = titulo.getattribute("titulo")
			end if
		end if 
		if ""&eltitulo <> "" then
			response.Write("<img src=../img/flecha.gif><font size=1>" & eltitulo & "</font>")
		end if
	end Sub

	'Poner los titulos de las paginas. (Se usa en el index.)
	Function titulohtml(objeto)
		Dim titulo, tituloPadre, salida
		salida = ""
		titulo = ""&objeto.getattribute("titulo")
		tituloPadre = ""&objeto.parentNode.getattribute("titulo")
	
		if titulo <> "" then
			salida = titulo
		else
			if tituloPadre <> "" then
				salida = tituloPadre
			end if
		end if
		titulohtml = salida
	end Function
		
	' Rutinas botoneras
	' --------------------------------------------
	Sub PintaBotoneras(raiz)
		if typeOK(raiz) then
			dim tipobotonera
			' Si tiene hijos se basa en su propio atributo
			if raiz.haschildnodes then
				PintaBotonera raiz.childnodes , raiz.getAttribute("botonera")
	
			' Si no tiene hijos y no es la raiz (secciones)
			elseif ""&raiz.parentNode.nodename <> "secciones" then
				PintaSubBotonera raiz.parentNode.childnodes , raiz.parentNode.getAttribute("botonera")
	
			end if
		end if
	end sub
	
	Sub PintaBotonerasMovil(raiz)
		if typeOK(raiz) then
			dim tipobotonera
			' Si tiene hijos se basa en su propio atributo
			if raiz.haschildnodes then
				PintaBotoneraMovil raiz.childnodes , raiz.getAttribute("botonera")
	
			' Si no tiene hijos y no es la raiz (secciones)
			elseif ""&raiz.parentNode.nodename <> "secciones" then
				PintaSubBotoneraMovil raiz.parentNode.childnodes , raiz.parentNode.getAttribute("botonera")
	
			end if
		end if
	end sub
	
	sub PintaBotonera(Nodos,tipo)
		Dim oNodo, altoflash, cadena, n_b
		%><tr><td><%
		colorfondo = Nodos.item(0).parentNode.getattribute("colorfondo") 
		cadena = ""
		n_b = 0
		For Each oNodo In Nodos
			if oNodo.nodeType = 1 and ""&oNodo.getattribute("ocultarpublico") <> "1" then 
				n_b = n_b+1
				if oNodo.parentNode.getattribute("activo")<>"" then
					cadena = cadena & "activo=" & oNodo.parentNode.getattribute("activo") &"&"
				end if
				cadena = cadena & "tituloboton"& n_b &"="& oNodo.getattribute("titulo") &"&enlaceboton"& n_b &"=index.asp?secc="& secc &"/"& oNodo.nodename &"&titulocorto"& n_b &"="& oNodo.nodename &"&"
			end if
		Next
		cadena = utf(cadena & "totalbotones="& n_b & "&colorfondo="& colorfondo &"&secc="& secc &"&")
		if tipo<>"" then
		if tipo = 1 and n_b>0 then
			flash "subbotonera.swf",cadena,anchobotonera1,altobotonera1,"#FFFFFF"
		elseif tipo = 2 and n_b>0 then
			flash "subbotonera2l.swf",cadena,anchobotonera2,altobotonera2,"#FFFFFF" 
		end if
		end if
		%></td></tr><%
	End sub
	
	sub PintaBotoneraMovil(Nodos,tipo)
		Dim oNodo, altoflash, cadena, n_b
			For Each oNodo In Nodos
			if oNodo.nodeType = 1 and ""&oNodo.getattribute("ocultarpublico") <> "1" then 
%>
				<a href=index.asp?secc=<%=secc%>/<%=oNodo.nodename%>><%Response.write(oNodo.getattribute("titulo"))%></a><br>
<%				
			end if
		Next
	End sub
	
	
	Sub PintaSubBotonera(Nodos,tipo)
		%><tr><td><%
		Dim oNodo, altoflash, cadena
	
		colorfondo = Nodos.item(0).parentNode.getattribute("colorfondo") 
		cadena = ""
		n_b = 0
		posicion = Instr(secc,Nodos.item(0).parentNode.nodename)+len(Nodos.item(0).parentNode.nodename)
		ruta = Left(secc,posicion)
		For Each oNodo In Nodos
			if oNodo.nodeType = 1 and ""&oNodo.getattribute("ocultarpublico") <> "1" then 
				n_b = n_b+1
				cadena = cadena & "tituloboton"& n_b &"="& oNodo.getattribute("titulo") &"&enlaceboton"& n_b &"=index.asp?secc="& ruta & oNodo.nodename & "&titulocorto"& n_b &"="& oNodo.nodename & "&"
			end if
		Next
		cadena = utf(cadena & "totalbotones="& n_b & "&colorfondo="&colorfondo &"&secc="& secc &"&")
		if ""&tipo = "1"  and n_b>0 then
			flash "subbotonera.swf",cadena,anchobotonera1,altobotonera1 ,"#FFFFFF"
		elseif ""&tipo = "2"  and n_b>0 then
			flash "subbotonera2l.swf",cadena,anchobotonera2,altobotonera2,"#FFFFFF"
		end if
		%></td></tr><%
	End sub
	
	
	
	Sub PintaSubBotoneraMovil(Nodos,tipo)

		Dim oNodo, altoflash, cadena
	
		posicion = Instr(secc,Nodos.item(0).parentNode.nodename)+len(Nodos.item(0).parentNode.nodename)
		ruta = Left(secc,posicion)
		For Each oNodo In Nodos
			if oNodo.nodeType = 1 and ""&oNodo.getattribute("ocultarpublico") <> "1" then 

				if strcomp(secc,ruta&oNodo.nodename)<>0 then
				%>				
				<a href=index.asp?secc=<%=ruta%><%=oNodo.nodename%>><%Response.write(oNodo.getattribute("titulo"))%></a><br>
				<%
				else
				%>
				* <%Response.write(oNodo.getattribute("titulo"))%><br>
				<%
				end if
			end if
		Next

	End sub

%>