<!--#include virtual="/datos/inc_config_gen.asp" -->
<%
ruta_xmlsecciones="/secciones.xml"

if ""&session("cualid")="edicionmovil" then
	nav_sistema = "movil"
else
	nav_sistema = "pc"
end if

%>
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<%
secc = request("secc")
if unerror then
	Response.Write "<b>Error</b><br>" & msgerror
else

if session("usuario") = "" then%>
	No está validado
<%else
	if not getPermiso("edicion",session("idioma")) then%>
		No tiene parmiso para acceder a esta zona
	<%else%>
<html>
<head>
<title>aSkipper</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.lineabajo {
	border-bottom-width: thin;
	border-bottom-style: solid;
	border-bottom-color: #726DBE;
}

.aSelectLista {
	background-color: #f5f5f5;
	border: thin solid #f5f5f5;
	height: 15px;
}
.linkSeccion {
	text-decoration: none;
}
.bodyAdmin {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 7.5pt;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #f5f5f5;
}
.overFilaSeccion{
	background-color:#0099FF;
	color:#FFFFFF;
}
.outFilaSeccion{
	background-color:#f5f5f5;
	color:;
}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
	function editar(secc){
		ventana("lista.asp?ac=editarseccion&secc="+secc,"Editar",320,150,"scrollbars=no")
	}	
//-->
</script>
</head>
<body bgcolor="#E0EFF8" link="#666699" vlink="#666699" alink="#FF6600" class="bodyAdmin" marginheight="0" marginwidth="0" leftmargin="0" rightmargin="0">

<%
			Function elemento(cadena,numero)
				Dim coincidencia, n, resto
				coincidencia = 0
				' paso las "/"
				for n=0 to numero-1
					coincidencia = Instr(coincidencia+1,cadena,"/")
				next
				resto = Instr(coincidencia+1,cadena,"/")
				if resto>0 then
					elemento=Mid(cadena,coincidencia+1,resto-coincidencia-1)
				else
					elemento="noencontrado"				
				end if
			end Function
			
function pintaruta(valor,elxml)
	Dim rutaxml, n, devuelto
	set rutaxml = elxml
	for n=1 to valor
		Set rutaxml = rutaxml.parentnode
		devuelto = rutaxml.nodename&"/"&devuelto
	next
	pintaruta=devuelto
end function
%>
<script language="javascript" type="text/javascript">
function borrarNodo(lugar,secc) {
	if (confirm("¿Seguro que desea borrar esta sección?")) {
		location.href="lista.asp?ac=mover&desp=<%=request("desp")%>&dir=eliminar&lugar="+lugar+"&secc="+secc
	}
}
function nuevoNodo(lugar,pos,secc) {
	location.href="lista.asp?ac=nuevo&desp=<%=request("desp")%>&pos="+pos+"&lugar="+lugar+"&secc="+secc
}
function nuevoNodoHijo(lugar,pos,secc) {
	location.href="lista.asp?ac=nuevo&desp=<%=request("desp")%>&pos="+pos+"&lugar="+lugar+"&secc="+secc+"&tipo=hijo"
}
function cambiarVisiblePublico() {
	if (arguments.length >1) {
		newEstado = arguments[0]
		nodo = arguments[1]
		newEstado = (newEstado)? 0 : 1
		nodo = (nodo=="")? "" : nodo
		location.href='lista.asp?ac=cambiarVisiblePublico&desp=<%=request("desp")%>&newEstado='+newEstado+"&nodo="+nodo
	}
}
function ventana(theURL,winName,ancho,alto,features) { 
	var winl = (screen.width - ancho) / 2;
	var wint = (screen.height - alto) / 2;
	var paramet=features+',top='+wint+',left='+winl+',width='+ancho+',height='+alto;
	var splashWin=window.open(theURL,winName,paramet);
    splashWin.focus();
}

function cambioLinea(newEstado,nodo) {
	location.href='lista.asp?ac=lineasbotonera&desp=<%=request("desp")%>&newEstado='+newEstado+"&nodo="+nodo+"&secc=<%=secc%>"
}
function overFilaSeccion(fila){
	fila.className = "overFilaSeccion"
}
function outFilaSeccion(fila){
	fila.className = "outFilaSeccion"
}
function clickFila(enlace){
	//alert(enlace)
	parent.frames['principal'].location.href=enlace
}
</script>

<%
Sub ver(Nodos,seccion)
	Dim n, pos, separa
	Dim oNodo, xmlfile, anterior, rutanterior, secc, vueltaSola, lugar, voyapintar
	Dim editable, administrable, fija, botonera, pmiso, phijos, valorrotulo, ocultarpublico, opt_botonera
	Dim subir, bajar
	pos = 1
	
	For Each oNodo In Nodos

		if oNodo.nodeType = 1 then
		
			if (inStr(oNodo.getAttribute("compatible"),nav_sistema)) or (""&nav_sistema = "pc" and ""&oNodo.getAttribute("compatible") = "") then

				xmlfile=oNodo.nodename&".xml"
				Set anterior = oNodo
				rutanterior=""
				for n=0 to separador
					rutanterior=anterior.nodename&"/"&rutanterior
					Set anterior=anterior.parentnode
				next
				
				secc="/"&pintaruta(separador,oNodo)&oNodo.nodename
				vueltaSola="/"&pintaruta(separador,oNodo)
				xmlfile="../../"&session("idioma")&carpeta_delsitio&secc&"/"&oNodo.nodename&".xml"
				lugar=pintaruta(separador,oNodo)&oNodo.nodename
	
				
				editable = ""&oNodo.getattribute("editable")
				administrable = ""&oNodo.getattribute("administrable")
				fija = ""&oNodo.getattribute("fija")
				botonera = ""&oNodo.getattribute("botonera")
				pmiso = getPermisoParaRuta("edicion","esp",session("usuario"),lugar)
				phijos = getPermisoHijos(session("usuario"),lugar)
				plantilla = ""&oNodo.getAttribute("plantilla")
				miTitulo = oNodo.getAttribute("titulo")
				
				if plantilla <> "" then
					enlace = "/" & c_s & "plantillas/" & plantilla & ".asp?secc="&secc
				else
					enlace = "/" & c_s & session("idioma") & carpeta_delsitio & secc & "/" & nombreArchivo(secc) & ".asp?secc="&secc
				end if
	
				'  "ocultarpublico" es un atributo en el xml secciones que indica si la pagina se muestra en la modo usuario
				' no afecta ni tiene relacion con el atributo "visible"
				ocultarpublico = cbool("0"&oNodo.getattribute("ocultarpublico"))
				d = chr(24)
				if editable="1" and pmiso then
					href_ini = "<a href='"& enlace &"' class='linkSeccion' target='principal'>"
					href_fin = "</a>"				
				end if
				%>
				<table width="100%" cellpadding="1" cellspacing="0" onMouseOver="overFilaSeccion(this)" onMouseOut="outFilaSeccion(this)">
					<tr>
					<td>&nbsp;</td>
					<%for g=0 to separador-1%>
						<td><img src="../../spacer.gif" width="10" height="10"></td>
					<%next%>
					<%if pmiso then%>
						<td>
							<%if not ocultarpublico then%>
								<a href="JavaScript:cambiarVisiblePublico(0,'<%=secc%>')"><img src="../images/marcado.gif" width="10" height="10" border="0" alt="Desactivar"></a>
							<%else%>
								<a href="JavaScript:cambiarVisiblePublico(1,'<%=secc%>')"><img src="../images/marcado_no.gif" border="0" alt="Activar"></a>
							<%end if%>
						</td>
					<%end if%>
					<td>·</td>
					<td width="100%" align="left" valign="middle" onClick="clickFila('<%=enlace%>')">
					<%if ocultarpublico then%>
						<font color="#999999"><%=miTitulo%></font>
					<%else%>
						<%=miTitulo%>
					<%end if%>
					</td>
					<td align="right" valign="middle"><a href="JavaScript:editar('<%=secc%>')" title=" Editar "><img src="../images/editar.gif" border="0"></a></td>
					</tr>
	</table>
		<%
			end if ' fin de compatibilidad
		end if

		If oNodo.hasChildNodes Then
			separador=separador+1
			ver oNodo.childNodes,oNodo.nodename
			separador=separador-1
		End If
		Response.Flush()
	Next
End Sub

function quitaracentos(cadena)
	cadena2 = ""&cadena
	if cadena2 <> "" then
		cadena2 = Lcase(Replace(cadena2," ",""))
		cadena2 = Replace(cadena2,"á","a")
		cadena2 = Replace(cadena2,"é","e")
		cadena2 = Replace(cadena2,"í","i")
		cadena2 = Replace(cadena2,"ó","o")
		cadena2 = Replace(cadena2,"ú","u")
		quitaracentos = cadena2
	else
		quitaracentos = cadena2
	end if
end function

function limpiaParaNombreNodo (texto)
	for n=1 to len(texto)
		c = Mid(texto,n,1)
		if asc(c) >= 97 and asc(c) <= 122 then
			salida = salida & c
		elseif asc(c) >= 48 and asc(c) <= 57 then
			salida = salida & c
		elseif c = "_" then
			salida = salida & c
		end if
	next
	limpiaParaNombreNodo = salida
end function


function crearNombreNodo()
	crearNombreNodo = "seccion0" & num
end function

select case request.QueryString("ac")

case "editarseccion"

	set xmlObj = CreateObject("MSXML2.DOMDocument")
	if not xmlObj.Load(Server.MapPath("/" & c_s & session("idioma") & ruta_xmlsecciones)) then
		unerror = true : msgerror ="No se ha encontrado el archivo XML o contiene algún error.<br>Archivo: "&Request.Servervariables("PATH_TRANSLATED")&"."
	end if

	if not unerror then
		xp = "/pagina/secciones" & secc
'		on error resume next
		set seccion = xmlObj.selectNodes(xp).item(0)
		npos = 1
		if not typeOK(seccion) then
			unerror = true : msgerror = "No se ha encontrado el nodo indicado."
		else
			set padre = seccion.parentNode
			for each a in padre.childNodes
				if ""&a.nodeName <> ""&seccion.nodeName then
					npos = npos + 1
				else
					pos = npos
				end if
			next
		end if
		on error goto 0
		if not typeOK(seccion) then
			unerror = true : msgerror = "No se ha encontrado la sección indicada."
		end if		
	end if
	
	if not unerror then
		titulo = seccion.getAttribute("titulo")
		mibotonera = seccion.getAttribute("botonera")
		botoneraPadre = seccion.parentNode.getAttribute("botonera")
		administrable = cbool("0" & seccion.getAttribute("administrable"))
		padreAdministrable = cbool("0" & seccion.parentNode.getAttribute("administrable"))
		tengoHijos = cbool("0" & seccion.childNodes.length)
		soyFija = cbool("0" & seccion.getAttribute("fija"))
	end if

	algunaOpcion = false
%>

	    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%"  border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td bgcolor="#FFFFFF">
                  <table width="100%"  border="0" cellspacing="0" cellpadding="1">
                    <tr>
                      <td><font size="2" color="#0066FF"><b><%=titulo%></b></font></td>
                      <td align="right"><font color="#999999">(<%=pos%>)</font></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
              <table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#FFCC66">
              <tr>
                <td><table width="100%" border="0" cellpadding="1" cellspacing="0">
                    <tr>
                      <td width="9"><img src="../../spacer.gif" width="1" height="1"></td>
                      <td valign="top" bgcolor="#FFFFCC">
					  <table width="100%" border="0" cellpadding="2" cellspacing="0">

						<%if administrable then
							algunaOpcion = true%>
                          <tr>
                            <td align="left" valign="middle"><a href="#" id="nuevohijo" onClick="nuevoNodoHijo('<%=secc%>','<%=pos%>','<%=secc%>')"><img src="img_admin/nuevo.gif" alt="Nueva" width="13" height="13" border="0"></a></td>
                            <td width="100%" align="left" valign="middle"><nobr>
                            <label for="nuevohijo">Nueva secci&oacute;n hija</label>
                            </nobr></td>
                          </tr>
						  <%end if%>


						<%if padreAdministrable then
							algunaOpcion = true%>
                          <tr>
                            <td align="left" valign="middle"><a href="#" id="nuevo" onClick="nuevoNodo('<%=secc%>','<%=pos%>','<%=secc%>')"><img src="img_admin/nuevo.gif" alt="Nueva" width="13" height="13" border="0"></a></td>
                            <td width="100%" align="left" valign="middle"><nobr>
                            <label for="nuevo">Nueva secci&oacute;n hermana</label>
                            </nobr></td>
                          </tr>
						  <%end if%>


                          <%if administrable then
							  algunaOpcion = true%>
						  <tr>
                            <td align="left" valign="middle">
							<%if mibotonera = "1" then%>
                                <a href="#" id="boto" onClick="cambioLinea('2','<%=secc%>')"><img src="img_admin/botonera1.gif" alt=" Pasar a botonera de dos lineas " border="0"></a>
                                <%elseif mibotonera = "2" then%>
                                <a href="#" id="boto" onClick="cambioLinea('','<%=secc%>')"><img src="img_admin/botonera2.gif" alt=" Pasar a botonera de una linea " border="0"></a>
                                <%else%>
                                <a href="#" id="boto" onClick="cambioLinea('1','<%=secc%>')"><img src="img_admin/botoneraVacio.gif" alt=" Quitar botonera " border="0"></a>
                                <%end if%></td>
                            <td width="100%" align="left" valign="middle"><nobr><label for="boto">
							<%if mibotonera = "1" then%>
                            Cambiar a dos lineas de botones
							<%elseif mibotonera = "2" then%>
                            Quitar botonera
							<%else%>
                            Cambiar a una linea de botones
							<%end if%>
                            &nbsp;</label></nobr></td>
                          </tr>
						  <%end if%>
						  
                          <tr>
                            <td align="left" valign="middle"><a id="subir" href="lista.asp?ac=mover&dir=subir&secc=<%=secc%>&desp=<%=request("desp")%>&pos=<%=pos-1%>"><img src="img_admin/arriba.gif" width="13" height="13" alt="Subir" border="0"></a></td>
                            <td width="100%" align="left" valign="middle"><nobr>
                            <label for="subir">Subir posici&oacute;n</label>
                            </nobr></td>
                          </tr>
                          <tr>
                            <td align="left" valign="middle"><a id="bajar" href="lista.asp?ac=mover&dir=bajar&secc=<%=secc%>&desp=<%=request("desp")%>&pos=<%=pos-1%>"><img src="img_admin/abajo.gif" width="13" height="13" alt="Bajar" border="0"></a></td>
                            <td width="100%" align="left" valign="middle"><nobr>
                              <label for="bajar">Bajar posici&oacute;n </label>
                              &nbsp;</nobr></td>
                          </tr>
						  
						  <%if not tengoHijos and not soyFija then
							  algunaOpcion = true%>
                          <tr>
                            <td align="left" valign="middle"><a id="eliminar" href="javascript:borrarNodo('<%=lugar%>','<%=secc%>')"><img src="img_admin/eliminar.gif" alt="Eliminar" width="13" height="13" border="0"></a></td>
                            <td width="100%" align="left" valign="middle"><nobr><label for="eliminar">Eliminar</label>&nbsp;</nobr></td>
                          </tr>
						  <%end if%>
						  
                      </table></td>
                    </tr>
                </table></td>
              </tr>
            </table>
			<%if not unerror and 1=2 then%>
				<br><br><br><br><br><br>
				Resumen:<br>
				<%
				set atts = seccion.attributes
				for each att in atts
					Response.Write att.nodeName & ": "& att.nodeValue &"<br>"
				next
			end if
			
			if unerror then%>
				<script>window.close()</script>
			<%end if%></td>
          </tr>
</table>
        <%case "elegirplantilla"%>

<script>
function eleccion(np) {
	parent.opener.f1.plantilla.options[np-1].selected = true
	window.close()	
}
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="2">
  <tr><td>
Elija el tipo de página que quiere insertar:
</td><tr></table>

<table width="100%" border="1" cellpadding="0" cellspacing="6" bgcolor="#FFFFFF">
<tr>
<%
carpetaplant=Server.MapPath("../../plantillas")
inicial=1
dim fso
set fso = Server.CreateObject("Scripting.FileSystemObject")
Set Upload = Server.CreateObject("Persits.Upload")
if fso.FolderExists(carpetaplant) then
	dim SORTBY_TYPE
	SORTBY_TYPE = ""
	Set Dir = Upload.Directory( carpetaplant&"/*.*", SORTBY_TYPE)
	n=1
	For Each item in Dir
		bFile = carpetaplant&"/"&item.FileName
		if Strcomp(right(item.FileName,4),".asp")=0 then
			ficheroenxml = carpetaplant&"/"&replace(item.FileName,".asp",".xml")
			if Upload.FileExists (ficheroenxml) then
				nombreplant = replace(item.FileName,".asp","")

			
				%>
			  <td align="center"><br>
			  <img src="../../plantillas/<%=nombreplant%>.gif"><br>
			  <input name="plantilla" type="button" class="botonAdmin" onClick="eleccion('<%=n%>')" value="<%=nombreplant%>"></td>
			<%
				n=n+1
				if n mod 5 = 1 then
				%>
  </tr>
 				 <tr>
			  <%
			end if
			end if
		end if
		'if Upload.FileExists (bFile) then
		'	response.write("<br> el nobmre: "&item.FileName)
		'end if
	Next
	
	
	dim modulo, p
	modulo = n mod 5
	if modulo <> 1 then
	for p=modulo to 6-modulo
		%>
		<td><nobr>  </nobr></td>
		<%
	next
	end if
end if
Set Upload = Nothing
Set fso = Nothing%>
</tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="2">
  <tr><td align="right" valign="bottom">
  <br>
  <input type="button" class="botonAdmin" onClick="window.close()" value="Cancelar">

</td><tr></table>
<%
case "insertarnuevo" ' ----------------------------------------------------------------------------- nuevo nodo

	' Tomamos el título escrito y lo limpiamos
	dim tituloRecibido, nombreNodo
	nombreNodo = ""
	tituloRecibido = ""

	tituloRecibido = ""&replace(lcase(""&request.Form("titulo")),"#","")
	filtroHtml(tituloRecibido)
	
	' Nombre del nodo
	nombreNodo = replace(tituloRecibido,"    "," ")
	nombreNodo = replace(nombreNodo,"   "," ")
	nombreNodo = replace(nombreNodo,"  "," ")
	nombreNodo = replace(nombreNodo," ","_")
	nombreNodo = quitaracentos(nombreNodo) ' Quitar acentos
	
	letras = "abcdefghijklmnopqrstuvwxyz"
	c = left(nombreNodo,1)
	dim fin : fin = false
	while inStr(letras,c)=0 or fin
		if nombreNodo = "" then
			fin = true
		else
			nombreNodo = ""& right(nombreNodo,len(nombreNodo)-1)
			c = left(nombreNodo,1)
		end if
	wend
	
	if tituloRecibido = "" then
		unerror = true : msgerror = "No se ha recibido un título para la nueva sección."
	end if
	if secc = "" then
		unerror = true : msgerror = "No se ha recibido la sección."
	end if

	carpetaArchivo = "/" & c_s & session("idioma") & secc
	pos = request.QueryString("pos")
	lugar = ""&request.QueryString("lugar")
	if lugar = "" then
		unerror = true : msgerror = "No se ha recibido la sección actual."
	end if
	if pos = "" then
		unerror = true : msgerror = "No se ha recibido la posición."
	end if

	dim laplantilla
	laplantilla = ""&request.Form("plantilla")
	if laplantilla = "" then
		unerror = true : msgerror = "No se ha recibido la plantilla que desea utilizar."
	end if

	' Cargo el XML de secciones
	if not unerror then
		set xmlObj = CreateObject("MSXML.DOMDocument")
		if not xmlObj.Load(Server.MapPath("/" & c_s & session("idioma") &ruta_xmlsecciones)) then
			unerror = true : msgerror = "No se ha logrado cargar el XML de secciones."
		end if
	end if
	
	' Seteo el nodo de la sección actual (y su padre)
	if not unerror then
		dim nodoSeccionActual
		set nodoSeccionActual = xmlObj.selectNodes("/pagina/secciones" & secc)
		if not typeOK(nodoSeccionActual) then
			unerror = true : msgerror = "No se ha encontrado la sección indicado (1)."
		else
			set nodoSeccionActual = nodoSeccionActual.item(0)
		end if
		if not typeOK(nodoSeccionActual) then
			unerror = true : msgerror = "No se ha encontrado la sección indicado (2)."
		else
			set padre = nodoSeccionActual.parentNode
		end if
	end if

	if request("tipo") = "hijo" then
		msecc = secc
	else
		msecc = replace(secc& "|","/"&nodoSeccionActual.nodeName&"|","")
	end if

	elNombreNodo = limpiaParaNombreNodo(nombreNodo)
	if elNombreNodo = "" then
		unerror = true : msgerror = "El nombre no es válido.<br>Por favor, introduzca una o más letras que no sean signos."
	end if
	
	if not unerror then
		' Comprobar que no existe un nodo con el mismo nombre
		set nodo_igual = xmlObj.selectSingleNode("pagina/secciones"& msecc & "/" &elNombreNodo)
		if typeOK(nodo_igual) then
			Response.Redirect("lista.asp?ac=nuevo&desp=&pos=1&lugar="& lugar &"&secc="& secc &"&tipo="& request("tipo") &"&titulo="& tituloRecibido &"&msgerror=Ya existe una página con ese título.")
		end if

		' CREAR CARPETA Y ARCHIVOS CORRESPONDIENTES
		rutaCarpetaNueva = "/" & c_s & session("idioma") & msecc & "/" & elNombreNodo
		if not nuevaCarpeta(server.MapPath(rutaCarpetaNueva),true) then
			unerror = true : msgerror = "No se ha logrado crear el directorio para la sección."
		end if
	end if

	' Copiamos los archivos de la plantilla a la carpeta de nuestra nueva sección	
	if not unerror then
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		set xmlPlantilla = fso.getFile(Server.MapPath("/" & c_s & "plantillas/"& laplantilla &".xml"))
		if typeOK(xmlPlantilla) then
			xmlPlantilla.copy(server.MapPath(rutaCarpetaNueva & "/" & elNombreNodo &".xml"))
		else
			unerror = true : msgerror = "No se ha encontrado el XML de la plantilla solicitada."
		end if
	end if
	
	' Definir atributos
	if not unerror then
		set nuevo = xmlObj.createElement(elNombreNodo)

		' Si padre del nuevo nodo que estamos creado es administrable, el hijo neuvo también será administrable.
		if ""&nodoSeccionActual.getAttribute("hijosadministrables") = "1" then
			set att = xmlObj.createAttribute("administrable")
			att.nodeValue = "1"
			nuevo.setAttributeNode(att)
			
			' Sus hijos tambien serán administrables
			set att = xmlObj.createAttribute("hijosadministrables")
			att.nodeValue = "1"
			nuevo.setAttributeNode(att)
		end if

		set att = xmlObj.createAttribute("titulo")
		att.nodeValue = tituloRecibido
		nuevo.setAttributeNode(att)

		' La "botonera" de hijo será igual a "botonera" del padre. :P
		set att = xmlObj.createAttribute("botonera")
		att.nodeValue = ""&padre.getAttribute("botonera")
		nuevo.setAttributeNode(att)

		set att = xmlObj.createAttribute("activo")
		att.nodeValue = "1"
		nuevo.setAttributeNode(att)
		
		set att = xmlObj.createAttribute("editable")
		att.nodeValue = "1"
		nuevo.setAttributeNode(att)

		set att = xmlObj.createAttribute("visible")
		att.nodeValue = "1"
		nuevo.setAttributeNode(att)

		set att = xmlObj.createAttribute("ocultarpublico")
		att.nodeValue="1"
		nuevo.setAttributeNode(att)

		set att = xmlObj.createAttribute("plantilla")
		att.nodeValue = laplantilla
		nuevo.setAttributeNode(att)

		set att = xmlObj.createAttribute("estilotitulo")
		att.nodeValue = "1"
		nuevo.setAttributeNode(att)
		
		set att = nothing

		' Insertamos hermano o hijos según hayamos escogido
		if request("tipo") = "hijo" then
			nodoSeccionActual.appendChild(nuevo)
		else
			padre.insertBefore nuevo, padre.childNodes.item(pos)
		end if
		xmlObj.save Server.MapPath("/" & c_s & session("idioma") & ruta_xmlsecciones)
	
	end if


	if not unerror then%>
		<script language="javascript" type="text/javascript">
		<!--
		try{
			parent.opener.location.href=parent.opener.location
		} catch(unerror){}
		window.close()
		//-->
		</script>
	<%else%>
	<table width="100%"  border="0" cellspacing="0" cellpadding="4">
  <tr>
    <td><font color=FF0000><%=msgerror%></font><br><br>
<input type="button" name="Volver" value="Volver" onClick="history.back()"></td>
  </tr>
</table>

		
	<%end if


case "nuevo" ' ---------------------
%>
<script>
	function validar(f) {
		if(f.titulo.value == ""){
			alert("Escriba el titulo de la nueva página.")
			f.titulo.focus()
			return false
		}
/*		VALIDADMOS CON ASP
		var patron = /\W/
		if(patron.test(f.titulo.value.replace(/\s|á|é|í|ó|ú|Á|É|Í|Ó|Ú|-/g))) {
			alert("Por favor, introduzca un título que contenga sólo letras, números y espacios.\nLa letra 'ñ' y el guión (-) no están pemitidas.")
			return false
		}
*/
		if (Number(f.titulo.value.charAt(0))){
			alert("El nombre de la sección no puede empezar con un número.")
			return false
		}
		return true
	}
	</script>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td>		
			<form action="lista.asp?ac=insertarnuevo&pos=<%=request.QueryString("pos")%>&lugar=<%=request.QueryString("lugar")%>&tipo=<%=request("tipo")%>" method="post" name="f1" onSubmit="return validar(this)">
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
		<tr bgcolor="#FFFFFF">
		  <td height="20" colspan="2">&nbsp;<span class="tituloazonaAdmin">Nueva secci&oacute;n</span>
		</tr>

		<%if ""&request.QueryString("msgerror") <> "" then%>
		<tr align="center">
		  <td colspan="2"><font color="#FF0000"><%=request.QueryString("msgerror")%></font></td>
		  </tr>
		  <%end if%>
		  
		<tr><td align="right">Título: <td><input name="titulo" type="text" value="<%=request("titulo")%>" size="20" maxlength="100">
		  <input type="hidden" name="secc" value="<%=request.QueryString("secc")%>">          
		</tr><tr>
		  <td align="right">
		      <a href="#" onClick="ventana('lista.asp?ac=elegirplantilla','a',700,450,'scrollbars=1')">Plantilla</a>: <br>
		  </td>
		  <td><select name="plantilla">
<%

dim carpetaplant, inicial, Upload, Dir, ficheroenxml, nombreplant
dim n, item, bfile
carpetaplant = Server.MapPath("../../plantillas")
inicial = 1
set fso = Server.CreateObject("Scripting.FileSystemObject")
Set Upload = Server.CreateObject("Persits.Upload")
if fso.FolderExists(carpetaplant) then
	Set Dir = Upload.Directory( carpetaplant&"/*.*", SORTBY_TYPE)
	n=0
	For Each item in Dir
		bFile = carpetaplant&"/"&item.FileName
		if Strcomp(right(item.FileName,4),".asp")=0 then
			ficheroenxml = carpetaplant&"/"&replace(item.FileName,".asp",".xml")
			if Upload.FileExists (ficheroenxml) then
				nombreplant = replace(item.FileName,".asp","")%>
            <option value="<%=nombreplant%>" <%if inicial=1 then response.write(" checked") : inicial=2 end if%>><%=nombreplant%></option>
            <%
				n=n+1
			end if
		end if
		'if Upload.FileExists (bFile) then
		'	response.write("<br> el nobmre: "&item.FileName)
		'end if
	Next

end if
Set Upload = Nothing
Set fso = Nothing%>
          </select></td>
		</tr>
		<tr>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  </tr>
		<tr>
		  <td align="right">&nbsp;</td>
		  <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
            <input name="submit" type="submit" class="botonAdmin" value="Aceptar"></td>
		  </tr>
	</table>
			</form><script>f1.titulo.focus()</script>
		</td>
	  </tr>
</table>

<%case "propiedades"
	
	idioma = session("idioma")

	secc = ""&request("secc")
	nombreactual = filtroHtml(request("nombreactual"))
	nombre = replace(request("nombre"),"#","")
	nombre = filtroHtml(nombre)
	estilotitulo = ""&request.Form("estilotitulo")
	seccionesxml = "/"& c_s & idioma & ruta_xmlsecciones
	

if nombre<>"" and nombreactual<>"" and secc <> "" then
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath(seccionesxml)) then
		set rNodo = xmlObj.selectSingleNode("pagina/secciones" & secc)
		if typeName(rNodo) = "Nothing" then
			unerror = true : msgerror = "No se ha encontrado la seccion actual. Puede que halla sido borrada."
		end if
	else
		unerror = true : msgerror = "No se ha podido cargar el XML de secciones."
	end if
	
	if not unerror then
		set attTitulo = xmlObj.createAttribute("titulo")
		rNodo.setAttributeNode(attTitulo)
		attTitulo.nodeValue = nombre
		
		set attEstiloTitulo = xmlObj.createAttribute("estilotitulo")
		rNodo.setAttributeNode(attEstiloTitulo)
		attEstiloTitulo.nodeValue = estilotitulo
	end if
	
	if not unerror then
		xmlObj.save Server.MapPath(seccionesxml)
		if err = 0 then
		else
			unerror = true : msgerror = "Se ha producido un error al intentar guardar el XML de secciones."
		end if
	end if


	' Meta
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

	' Meta
	if not unerror then
		set ntitulo = nodoContenido.selectSingleNode("titulo")
		if typeOK(ntitulo) then
			titulo_largo = ntitulo.text
		end if

		set nkeys = nodoContenido.selectSingleNode("keys")
		if typeOK(nkeys) then
			keys = nkeys.text
		end if

		set ndescripcion = nodoContenido.selectSingleNode("descripcion")
		if typeOK(ndescripcion) then
			descripcion = ndescripcion.text
		end if

		if not typeOK(ntitulo) then
			set ntitulo = xmlObj.createElement("titulo")
			nodoContenido.appendChild(ntitulo)
		end if
		ntitulo.text = filtroHtml(request.Form("titulo_largo"))

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
	end if
	
	
	if not unerror then
		%><script language="javascript" type="text/javascript">
			<!--
			parent.opener.location.href = parent.opener.location
			window.close()
			//-->
		</script>
		<%
	else
		Response.Write(msgerror)
	end if
else

	' XML lectura
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
		if typeOK(ntitulo) then
			titulo_largo = ntitulo.text
		end if

		set nkeys = nodoContenido.selectSingleNode("keys")
		if typeOK(nkeys) then
			keys = nkeys.text
		end if

		set ndescripcion = nodoContenido.selectSingleNode("descripcion")
		if typeOK(ndescripcion) then
			descripcion = ndescripcion.text
		end if
	end if
	
%>

<form action="lista.asp?ac=propiedades" method="post" name="f" id="f" onSubmit="return validar(this)">
<script>
function validar(f) {
//	f.nombre.value = f.nombre.value.replace(/'/g,"")
	var nombreactual = "<%=request("nombreactual")%>"
	if (f.nombre.value == "") {
		alert("Escriba el nuevo título para la sección \"<%=request("nombreactual")%>\".");
		f.nombre.focus()
		return false
	}
	return true
}
</script>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td height="20">&nbsp;<span class="tituloazonaAdmin">Cambiar t&iacute;tulo de secci&oacute;n</span></td>
    </tr>
  <tr align="center">
    <td><table width="400" border="0" cellspacing="0" cellpadding="10">
      <tr>
        <td><fieldset>
          <legend>T&iacute;tulo</legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="0">
            <tr valign="top">
              <td align="right">T&iacute;tulo p&aacute;gina: </td>
              <td><input name="titulo_largo" type="text" class="campoAdmin" id="titulo_largo" value="<%=titulo_largo%>" size="40" maxlength="255"></td>
            </tr>
            <tr valign="top">
              <td align="right">T&iacute;tulo secci&oacute;n:</td>
              <td><input name="nombre" type="text" class="campoAdmin" value="<%=request("nombreactual")%>" size="40" maxlength="30">
                  <input name="nombreactual" type="hidden" id="nombreactual" value="<%=request("nombreactual")%>">
                  <input name="secc" type="hidden" id="rutanodo" value="<%=secc%>">
              </td>
            </tr>
            <tr valign="top">
              <td align="right">Estilo de t&iacute;tulo: </td>
              <td><select name="estilotitulo" class="campoAdmin" id="estilotitulo">
                  <option value="0"<%if ""&request.QueryString("estilotitulo")="0" then%> selected<%end if%>>Ninguno</option>
                  <option value="1"<%if ""&request.QueryString("estilotitulo")="1" then%> selected<%end if%>>Normal</option>
              </select></td>
            </tr>
          </table>
        </fieldset></td>
      </tr>
    </table>

      <table width="380"  border="0" cellpadding="1" cellspacing="0" bgcolor="#FF0000">
        <tr>
          <td align="center"><table width="100%"  border="0" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td><font color="#FF0000" size="1">AVISO:</font><font size="1"> la URL para acceder a esta p&aacute;gina no cambiar&aacute;. <a href="javascript:alert('\nLa URL para acceder a esta p&aacute;gina no cambiar&aacute;\n\nSi el contenido de la secci&oacute;n ser&aacute; modificado sensiblemenrte, le recomendamos que borre esta secci&oacute;n y cree una nueva.')">[+] </a></font></td>
              </tr>
          </table></td>
        </tr>
      </table>



        <table width="400" border="0" cellspacing="0" cellpadding="10">
          <tr>
            <td><fieldset>
            <legend>Meta-tags</legend>
            <table width="100%" border="0" cellpadding="3" cellspacing="0">
              <tr valign="top">
                <td align="right">Palabras clave:</td>
                <td><input name="keys" type="text" class="campoAdmin" value="<%=keys%>" size="40" maxlength="200"></td>
              </tr>
              <tr valign="top">
                <td align="right">Descripci&oacute;n:</td>
                <td><textarea name="descripcion" cols="39" rows="5" wrap="virtual" class="areaAdmin"><%=descripcion%></textarea></td>
              </tr>
            </table>
            </fieldset></td>
          </tr>
        </table>      
        <table width="400" border="0" cellspacing="0" cellpadding="10">
          <tr>
            <td><fieldset>
              <legend>Direcciones a esta p&aacute;gina</legend>
              <table width="100%" border="0" cellpadding="3" cellspacing="0">
                <tr valign="top">
                  <td align="right">Enlace interno:</td>
                  <td><input type="text" class="campoAdmin" value="index.asp?secc=<%=secc%>" size="40" maxlength="200"></td>
                </tr>
                <tr valign="top">
                  <td align="right">Enlace externo:</td>
                  <td><input type="text" class="campoAdmin" value="<%=request.ServerVariables("HTTP_HOST")%>/<%=idioma%>/index.asp?secc=<%=secc%>" size="40" maxlength="200"></td>
                </tr>
              </table>
            </fieldset></td>
          </tr>
        </table>
        <table width="400" border="0" cellspacing="0" cellpadding="10">
          <tr>
            <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.close()" value="Cancelar">
              <input name="" type="submit" class="botonAdmin" value="Enviar">            </td>
          </tr>
        </table></td>
    </tr>
</table>
</form>
<br>
<br>
<%
end if


 ' ------------------------------------------------------------------------------- subir, bajar y eliminar
case "mover"
if secc <> "" then
	secc = right(secc,len(secc)-1)
	dir = request.QueryString("dir")
	pos = request.QueryString("pos")
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath("../..")&"/"&session("idioma")&ruta_xmlsecciones) then
		set rNodo = xmlObj.selectSingleNode("pagina/secciones").selectSingleNode(secc)
		if typeName(rNodo) <> "Nothing" then
			set padre = rNodo.parentNode
			
			select case dir
			case "eliminar"
				padre.removeChild(rNodo)
				xmlObj.save Server.MapPath("/" & c_s & session("idioma") & ruta_xmlsecciones)
				' Borrar carpeta y archivos
				carpeta = server.MapPath("/" & c_s & session("idioma") & "/" &  secc)

				set fso = Server.CreateObject("Scripting.FileSystemObject")

				Set Upload = Server.CreateObject("Persits.Upload")

				if fso.FolderExists(carpeta) then
					Set Dir = Upload.Directory( carpeta&"/*.*", SORTBY_TYPE)
					For Each item in Dir
						bFile = carpeta&"/"&item.FileName
						if Upload.FileExists (bFile) then
							Upload.DeleteFile bFile
						end if
					Next
					Upload.RemoveDirectory carpeta
				end if
				Set Upload = Nothing
				Set fso = Nothing
				%>
				<script language="javascript" type="text/javascript">
				parent.opener.location.href=parent.opener.location</script>
				<%
			dim bNodo
			case "subir"
				if pos <> "" then
					set bNodo = padre.removeChild(rNodo)
					padre.insertBefore bNodo, padre.childNodes.item(pos-1)
					set bNodo = Nothing
					xmlObj.save Server.MapPath("/" & c_s & session("idioma") & ruta_xmlsecciones)
				else
					%><script>alert("Error 3. Faltan datos.")</script><%
				end if
			case "bajar"
				if pos <> "" then
					set bNodo = padre.removeChild(rNodo)
					padre.insertBefore bNodo, padre.childNodes.item(int(pos)+1)
					xmlObj.save Server.MapPath("/" & c_s & session("idioma") & ruta_xmlsecciones)
					set bNodo = Nothing
				else
					%><script>alert("Error 3. Faltan datos.")</script><%
				end if
			case else
				%><script>alert("Error 2. Faltan datos.")</script><%
			end select
			set rNodo = Nothing
			set padre = Nothing
			if not unerror then
				'Response.Redirect("lista.asp?desp="&request("desp"))
				%>
				<script language="javascript" type="text/javascript">
				<!--
				parent.opener.location = parent.opener.location
				location.href="lista.asp?ac=editarseccion&secc=/<%=secc%>"
				//-->
				</script>
				<%
			end if
			'-------
		end if
	end if
else
	%><script>alert("Error 1. Faltan datos")</script><%
'	Response.Redirect("lista.asp")
end if

case "cambiarVisiblePublico"

dim newEstado, nodo, lugar, att
newEstado = ""&request.QueryString("newEstado")
nodo = ""&request.QueryString("nodo")
if newEstado <> "" and nodo <> "" then
	nodo = replace(nodo,"\","/")
	nodo = replace(nodo,"../","")
	if mid(nodo,len(nodo)) = "/" then
		nodo = left(nodo,len(nodo)-1)
	end if
else
	unerror = true : msgerror = "No se han recibido todos los datos necesarios."
end if

if not unerror then
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath("../../"&session("idioma")&ruta_xmlsecciones)) then
		set lugar = xmlObj.selectSingleNode("pagina/secciones"&nodo)
		if typename(lugar) = "Nothing" then
			unerror = true : msgerror = "No se ha encontrado el nodo indicado."
		end if
	end if
end if

if not unerror then
	set att = xmlObj.createAttribute("ocultarpublico")
	lugar.setAttributeNode(att)
	att.nodeValue = newEstado
	if err=0 then
		xmlObj.save Server.MapPath("../../"&session("idioma")&ruta_xmlsecciones)
	Set att=Nothing		
		Set xmlObj=Nothing		
		if err<>0 then
			unerror = true : msgerror = "No se ha podido guardar el XML."
		end if
	else
		unerror = true : msgerror = "Se ha producido un error al crear el nuevo estado."
	end if
end if

if not unerror then
	Response.Redirect("lista.asp?desp="&request("desp"))
else
	Response.Redirect("lista.asp?msg="&msgerror)
end if

case "lineasbotonera"

newEstado = ""&request.QueryString("newEstado")
nodo = ""&request.QueryString("nodo")
if nodo = "" then
	unerror = true : msgerror = "No se han recibido todos los datos necesarios."
end if

if not unerror then
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath("/" & c_s & session("idioma") & ruta_xmlsecciones)) then
		on error resume next
		set lugar = xmlObj.selectSingleNode("pagina/secciones"&nodo)
		on error goto 0
		if not typeOK(lugar) then
			unerror = true : msgerror = "No se ha encontrado el nodo indicado."
		end if
	end if
end if

if not unerror then

'	' Le asignamos el mismo valor de Botonera a los hijos
'	dim nodosdelugar, pe
'	Set nodosdelugar=lugar.childnodes
'	for pe=0 to nodosdelugar.length-1
'		set att = xmlObj.createAttribute("botonera")
'		nodosdelugar.item(pe).setAttributeNode(att)
'		att.nodeValue = newEstado
'		Set att=Nothing		
'	next

	' Cambio el valor de "botonera" para el nodo seleccionado
	set att = xmlObj.createAttribute("botonera")
	lugar.setAttributeNode(att)
	att.nodeValue = newEstado
	if err=0 then
		xmlObj.save Server.MapPath("/" & c_s & session("idioma") & ruta_xmlsecciones)
		Set att = Nothing		
		Set xmlObj = Nothing
		if err<>0 then
			unerror = true : msgerror = "No se ha podido guardar el XML."
		end if
	else
		unerror = true : msgerror = "Se ha producido un error al crear el nuevo estado."
	end if

	
end if

if not unerror then
	Response.Redirect("lista.asp?ac=editarseccion&secc="& secc &"&desp="&request("desp"))
else
	Response.Redirect("lista.asp?msg="&msgerror)
end if
 
case else%>
<br>

			<%Dim valorant, separador, xmlObj
			separador = 0
			Set xmlObj = CreateObject("MSXML.DOMDocument")
			if xmlObj.Load(Server.MapPath("/" & c_s & session("idioma")&ruta_xmlsecciones)) then
				ver xmlObj.selectsinglenode("/pagina/secciones").childnodes,"principal"
			Else
				response.Write("No se ha logrado cargar el XML en la ruta: " & Server.MapPath("/" & c_s & session("idioma")&ruta_xmlsecciones))
			End If
			set fso = Nothing%>

<%end select

' Refresca al recibir 'rutavueltaR'
dim rutavueltaR
rutavueltaR = request.QueryString("rutavueltaR")
if rutavueltaR <> "" then
	%><script>parent.frames['principal'].location.href=parent.frames['principal'].location</script><%
end if

if request.QueryString("msg")<>""then%>
	<script>alert("<%=request.QueryString("msg")%>")</script>
<%end if%>
</body>
</html>
		<%end if
	end if
end if%>