<!--#include virtual="/admin/inc_rutinas.asp" -->
<%

	str_conn_usuarios = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("/" & c_s & "datos/esp/usuarios/usuarios.mdb")
	public conn_activa_usuarios
	set conn_activa_usuarios = server.CreateObject("ADODB.Connection")
	on error resume next
	conn_activa_usuarios.Open str_conn_usuarios
	if err<> 0 then
		unerror = true : msgerror = "No se ha encontrado la base de datos de usuarios."
	end if
	on error goto 0
	
Dim usu
ruta_xml_grupos = "/" & c_s & "datos/grupos.xml"
ruta_xml_admin_config = "/" & c_s & "datos/xml_admin_config.xml"

' ARCHIVO XML
set xml_grupos = CreateObject("MSXML2.DOMDocument")
if not xml_grupos.Load(Server.MapPath(ruta_xml_grupos)) then
	unerror = true : msgerror ="No se ha encontrado el archivo XML o contiene algún error.<br>Archivo: "&Request.Servervariables("PATH_TRANSLATED")&"."
else
	' NODO: nodo_grupos
	Dim nodo_grupos
	set nodo_grupos = xml_grupos.selectSingleNode("datos/grupos")
	if not typeOK(nodo_grupos) then
		unerror = true : msgerror = "No se ha encontrado el nodo [nodo_grupos]"
	else
		' NODO: GRUPOS
		Dim grupos
		set grupos = xml_grupos.selectSingleNode("datos/grupos")
		if not typeOK(grupos) then
			unerror = true : msgerror = "No se ha encontrado el nodo [grupos]"
		end if
	end if
end if

' Cargo el XML de configuración
' ARCHIVO XML
Dim xmlConfig
set xmlConfig = CreateObject("MSXML.DOMDocument")
if not xmlConfig.Load(Server.MapPath(ruta_xml_admin_config)) then
	unerror = true : msgerror ="No se ha encontrado el archivo XML de configuración.<br>Archivo: "&Request.Servervariables("PATH_TRANSLATED")&".<br>Ruta:"&Server.MapPath(ruta_xml_admin_config)&""
else
	' NODO: cualidades
	Dim cualidades
	set cualidades = xmlConfig.selectSingleNode("configuracion")
	if not typeOK(cualidades) then
		unerror = true : msgerror = "No se ha encontrado el nodo [configuracion]"
	end if
end if


'
' Parametro indicado en cada página que incluye este archivo que nos indica si para ver esta página es necesario estar validado
dim usuariologeado
if usuariologeado and numero(session("usuario")) = 0 then
	unerror = true : msgerror = "Usted no est&aacute; validado.<br>Puede que su sesi&oacute;n de usuario haya expirado por estar demasiado tiempo inactiva.<br>Por favor, val&iacute;dese para acceder al contenido." : coderror = 11
end if

'
' Comprobar cualidad
function getCualidad(cualidad)
	if ""&cualidad <> "" then
		for each cualid in cualidades.childNodes
			if ""&cualidad = ""&cualid.nodeName then
				getCualidad = true
				exit function
			end if
		next
	end if
	getCualidad = false
end function

'
' Obtener el "título" de la cualidad indicada
function getTituloCualidad(cualidad)
	if ""&cualidad <> "" then
		for each cualid in cualidades.childNodes
			if ""&cualidad = ""&cualid.nodeName then
				getTituloCualidad = cualid.getAttribute("nombre")
				exit function
			end if
		next
	end if
	getTituloCualidad = ""
end function




'
' Setear el nodo grupo indicado en el parametro codigo de grupo
function setGrupo(setGrupo_codigo)
	dim codigo : codigo = ""&setGrupo_codigo
	for each grupo in grupos.childNodes
		if ""&grupo.getAttribute("id") = codigo then
			set setGrupo = grupo
			exit function
		end if
	next
	set setGrupo = nothing
end function

'
' Función para crear los campos de formularios según el XML
sub campoForm(usuario)
	Response.Write "incluida"
	if typeOK(usuario) then
		select case usuario.getAttribute("tipo")
			case "texto"%>
				<input name="<%=usuario.getAttribute("nombrecorto")%>" type="text" size="<%=usuario.getAttribute("ancho")%>" maxlength="<%=usuario.getAttribute("max")%>" value="<%=request.Form(usuario.getAttribute("nombrecorto"))%>" class="campoAdmin">
			<%case "email"%>
				<input name="<%=usuario.getAttribute("nombrecorto")%>" type="text" size="<%=usuario.getAttribute("ancho")%>" maxlength="<%=usuario.getAttribute("max")%>" value="<%=request.Form(usuario.getAttribute("nombrecorto"))%>" class="campoAdmin">
			<%case "dni"%>
				<input name="<%=usuario.getAttribute("nombrecorto")%>" type="text" size="<%=usuario.getAttribute("ancho")%>" maxlength="<%=usuario.getAttribute("max")%>" value="<%=request.Form(usuario.getAttribute("nombrecorto"))%>" class="campoAdmin">
			<%case "desplegable"%>
				<select name="<%=usuario.getAttribute("nombrecorto")%>" class="campoAdmin">
				<%for n=1 to usuario.childnodes.length%>
					<option value="<%=usuario.childnodes.item(n-1).text%>" <% if usuario.childnodes.item(n-1).getattribute("seleccionado")=1 then %>selected<%end if%>><%=usuario.childnodes.item(n-1).text%></option>
				<%next%>
				</select>
			<%case "multiselect"%>
				<select name="<%=usuario.getAttribute("nombrecorto")%>" size="<%=usuario.getAttribute("alto")%>" multiple class="campoAdmin">
				<%for n=1 to usuario.childnodes.length%>
					<option value="<%=usuario.childnodes.item(n-1).text%>" <% if usuario.childnodes.item(n-1).getattribute("seleccionado")=1 then %>selected<%end if%>><%=usuario.childnodes.item(n-1).text%></option>
				<%next%>
				</select>
		<%end select
	end if
end sub


sub campoFormLleno(usuario,codigo)
	Set usuariotem=usuario
	select case usuariotem.getAttribute("tipo")
		case "texto"
			%><input name="<%=usuariotem.getAttribute("nombrecorto")%>" type="text" size="<%=usuariotem.getAttribute("ancho")%>" maxlength="<%=usuariotem.getAttribute("max")%>" value="<%=getDatoUsu(codigo,usuariotem.getAttribute("nombrecorto"))%>" class="campoAdmin"><%
		case "email"
			%><input name="<%=usuariotem.getAttribute("nombrecorto")%>" type="text" size="<%=usuariotem.getAttribute("ancho")%>" maxlength="<%=usuariotem.getAttribute("max")%>" value="<%=getDatoUsu(codigo,usuariotem.getAttribute("nombrecorto"))%>" class="campoAdmin"><%
		case "dni"
			%><input name="<%=usuariotem.getAttribute("nombrecorto")%>" type="text" size="<%=usuariotem.getAttribute("ancho")%>" maxlength="<%=usuariotem.getAttribute("max")%>" value="<%=getDatoUsu(codigo,usuariotem.getAttribute("nombrecorto"))%>" class="campoAdmin"><%
		case "desplegable"
			pordefecto = getDatoUsu(codigo,usuariotem.getAttribute("nombrecorto"))%>
			<select name="<%=usuariotem.getAttribute("nombrecorto")%>" class="campoAdmin">
				<%for n=1 to usuariotem.childnodes.length%>
					<option value="<%=usuariotem.childnodes.item(n-1).text%>" <% if usuariotem.childnodes.item(n-1).text=pordefecto then %>selected<%end if%>><%=usuariotem.childnodes.item(n-1).text%></option>
				<%next%>
			</select>
		<%case "multiselect"
			pordefecto = getDatoUsu(codigo,usuariotem.getAttribute("nombrecorto"))%>
			<select name="<%=usuariotem.getAttribute("nombrecorto")%>" size="<%=usuariotem.getAttribute("alto")%>" multiple class="campoAdmin">
				<%for n=1 to usuariotem.childnodes.length%>
					<option value="<%=usuariotem.childnodes.item(n-1).text%>" <% if estaEnCadena(getDatoUsu(codigo,"marca"),usuariotem.childnodes.item(n-1).text) then %>selected<%end if%>><%=usuariotem.childnodes.item(n-1).text%></option>
				<%next%>
			</select>
		<%case "opcion"
			pordefecto = getDatoUsu(codigo,usuariotem.getAttribute("nombrecorto"))
			for n=1 to usuariotem.childnodes.length%>
				<input type="radio" name="<%=usuariotem.getAttribute("nombrecorto")%>" id="<%=usuariotem.getAttribute("nombrecorto") & n%>" value="<%=usuariotem.childnodes.item(n-1).getAttribute("valor")%>"<%if usuariotem.childnodes.item(n-1).getAttribute("valor") = pordefecto then %> checked<%end if%>><label for="<%=usuariotem.getAttribute("nombrecorto") & n%>"><%=usuariotem.childnodes.item(n-1).getAttribute("titulo")%></label>
			<%next
		end select
	Set usuariotem=Nothing
end sub



'
' Validar Email
function validarDni(dni)
	dnit = trim(dni)
	if isNumeric(dni) and len(dni) = 8 then
		validarDni = true
	else
		validarDni = false
	end if
end function


'
' Funcion para validar los formatos de los email
function validarEmail(email) 
    dim partes, parte, i, c 
    'rompo el email en dos partes, antes y después de la arroba 
    partes = Split(email, "@") 
    if UBound(partes) <> 1 then 
       'si el mayor indice del array es distinto de 1 es que no he obtenido las dos partes 
       validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela." 
       exit function 
    end if 
    'para cada parte, compruebo varias cosas 
    for each parte in partes 
       'Compruebo que tiene algún caracter 
       if Len(parte) <= 0 then 
          validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela."  
          exit function 
       end if 
       'para cada caracter de la parte 
       for i = 1 to Len(parte) 
          'tomo el caracter actual 
          c = Lcase(Mid(parte, i, 1)) 
          'miro a ver si ese caracter es uno de los permitidos 
          if InStr("._-abcdefghijklmnopqrstuvwxyz", c) <= 0 and not IsNumeric(c) then 
             validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela."  
             exit function 
          end if 
       next 
       'si la parte actual acaba o empieza en punto la dirección no es válida 
       if Left(parte, 1) = "." or Right(parte, 1) = "." then 
          validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela."  
          exit function 
       end if 
    next 
    'si en la segunda parte del email no tenemos un punto es que va mal 
    if InStr(partes(1), ".") <= 0 then 
       validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela." 
       exit function 
    end if 
    'calculo cuantos caracteres hay después del último punto de la segunda parte del mail 
    i = Len(partes(1)) - InStrRev(partes(1), ".") 
    'si el número de caracteres es distinto de 2 y 3 
'    if not (i = 2 or i = 3) then 
'       validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela." 
'       exit function 
'    end if 
    'si encuentro dos puntos seguidos tampoco va bien 
    if InStr(email, "..") > 0 then 
       validarEmail="Una dirección e-mail no puede contener dos espacios seguidos. Por favor, revíselo." 
       exit function 
    end if 
    validarEmail = true 
end function 

'
' Convierte cualquier texto en un nombre válido para ser nombre de usuario Skipper
function toNombreUsuario (nombre)
	dim letras
	letras = "abcdefghijklmnopqrstuvwxyz0123456789"
	c_ok = ""
	nombre = replace(lcase(""&nombre),"ñ","n")
	nombre = quitarAcentos(nombre)
	for n=1 to len(nombre)
		c = mid(nombre,n,1)
		if inStr(letras,c) > 0 then
			c_ok = c_ok & c
		end if
	next
	toNombreUsuario = c_ok
end function

'
' Validar un nombre de usuario, devuelvre 'true o la lista errores cometidos.
function validarNombreUsuario(param)
	dim nombre, salida
	nombre = lcase(param)
	salida = ""
	if inStr(nombre," ") >0 then
		salida = salida & "<br>No puede contener espacios."
	end if

	' Comprobar acentos
	Dim acentos
	acentos = "áéíóú"
	for n=1 to len(nombre)
		c = mid(nombre,n,1)
		if inStr(acentos,c) then
			conacentos = true
		end if
	next
	if conacentos then
		salida = salida & "<br>No puede contener acentos."
	end if

	' Comprobar signos
	Dim caracteres
	caracteres = "abcdefghijklmnopqrstuvwxyz"
	caracteres = caracteres & "0123456789"
	caracteres = caracteres & "@"
'	caracteres = caracteres & "ñ"
	caracteres = caracteres & "-"
	caracteres = caracteres & "_"
	for n=1 to len(nombre)
		c = mid(nombre,n,1)
		if inStr(caracteres,c) = 0 and inStr(listaCaracteres,c) = 0 and c <> " " then
			listaCaracteres = listaCaracteres & "'"& c &"' "
			errorCaracteres = true
		end if
	next
	if errorCaracteres then
		salida = salida & "<br>No puede contener los caracteres "& listaCaracteres
	end if

	if salida = "" then
		validarNombreUsuario = true
	else
		validarNombreUsuario = salida
	end if
end function


'
' Validar la clave de usuario, devuelvre 'true o la lista errores cometidos.
function validarClaveUsuario(param)
	dim clave, salida
	clave = lcase(param)
	salida = ""
	if inStr(clave," ") >0 then
		salida = salida & "<br>No puede contener espacios."
	end if

	' Comprobar acentos
	Dim acentos
	acentos = "áéíóú"
	for n=1 to len(clave)
		c = mid(clave,n,1)
		if inStr(acentos,c) then
			conacentos = true
		end if
	next
	if conacentos then
		salida = salida & "<br>No puede contener acentos."
	end if

	' Comprobar signos
	Dim caracteres
	caracteres = "abcdefghijklmnopqrstuvwxyz"
	caracteres = caracteres & "0123456789"
	caracteres = caracteres & "@"
'	caracteres = caracteres & "ñ"
	caracteres = caracteres & "-"
	caracteres = caracteres & "_"
	for n=1 to len(clave)
		c = mid(clave,n,1)
		if inStr(caracteres,c) = 0 and inStr(listaCaracteres,c) = 0 and c <> " " then
			listaCaracteres = listaCaracteres & "'"& c &"' "
			errorCaracteres = true
		end if
	next
	if errorCaracteres then
		salida = salida & "<br>No puede contener los caracteres "& listaCaracteres
	end if

	if salida = "" then
		validarClaveUsuario = true
	else
		validarClaveUsuario = salida
	end if
end function

'
' Devuelve el código de usuario a partir de su nombre, devuelve "" si no existe
function getCodigo(nombre_usuario)
	if ""&typeName(conn_activa_usuarios) = "Connection" then
		dim re, sql, codigo
		codigo = 0
		sql = "SELECT R_ID, R_TITULO FROM REGISTROS WHERE R_TITULO = '"& replace(nombre_usuario,"''","'") &"'"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_activa_usuarios
		re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
		re.Open()
		
		if not re.eof and not re.bof then
			codigo = re("R_ID")
		end if
	
		re.Close()
		set re = nothing
	end if
	getCodigo = codigo
end function

'
' Devuelve la clave de usuario a partir de su nombre, devuelve "" si no existe
function getClave(nombre_usuario)
	dim re, sql, clave
	clave = ""
	if ""&typeName(conn_activa_usuarios) = "Connection" then
		sql = "SELECT R_CLAVE, R_TITULO FROM REGISTROS WHERE R_TITULO = '"& replace(nombre_usuario,"''","'") &"'"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_activa_usuarios
		re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
		re.Open()

		if not re.eof and not re.bof then
			clave = re("R_CLAVE")
		end if
	
		re.close()
		set re = nothing
	end if
	getClave = clave
end function

'
' Devuelve la clave de usuario a partir de su nombre, devuelve "" si no existe
function getIdioma(nombre_usuario)
	if ""&typeName(conn_activa_usuarios) = "Connection" then
		dim re, sql, idioma
		idioma = ""
		sql = "SELECT R_IDIOMA, R_TITULO FROM REGISTROS WHERE R_TITULO = '"& replace(nombre_usuario,"''","'") &"'"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_activa_usuarios
		re.Source = sql
		re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
		re.Open()
		
		if not re.eof and not re.bof then
			idioma = re("R_IDIOMA")
		end if
	
		re.close()
		set re = nothing
	end if
	getIdioma = idioma
end function

'
function getPermisoGrupo(codigo_usuario,grupo)
	getPermisoGrupo = getGrupo(numero(codigo_usuario)) = numero(grupo)
end function

'
' Obtener el codigo de grupo del usuario cullo codigo le pasamos
function getGrupo(codigo_usuario)
	if ""&typeName(conn_activa_usuarios) = "Connection" then
		dim re, sql, grupo
		grupo = 0
		sql = "SELECT R_SECCION, R_ID FROM REGISTROS WHERE R_ID = "& numero(codigo_usuario)
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_activa_usuarios
		re.Source = sql
		re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
		re.Open()
		
		if not re.eof and not re.bof then
			grupo = re("R_SECCION")
		end if
	
		re.close()
		set re = nothing
	end if
	getGrupo = grupo
end function

'
' Función que nos devuelve el nombre de usuario correspondiente a la id de usuario que le pasamos
function getNombreUsuario(codigo_usuario)
	dim re, sql, nombre
	nombre = ""
	if ""&typeName(conn_activa_usuarios) = "Connection" then
		sql = "SELECT R_ID, R_TITULO FROM REGISTROS WHERE R_ID = "& numero(codigo_usuario)
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_activa_usuarios
		re.Source = sql
		re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
		re.Open()
		
		if not re.eof and not re.bof then
			nombre = re("R_TITULO")
		end if
	
		re.close()
		set re = nothing

	end if
	getNombreUsuario = nombre
end function

'
' Función que nos devuelve el nombre de usuario correspondiente a la id de usuario que le pasamos
function getEmailUsuario(codigo_usuario)
	dim re, sql, nombre
	nombre = ""
	if ""&typeName(conn_activa_usuarios) = "Connection" then
		sql = "SELECT R_ID, R_EMAIL FROM REGISTROS WHERE R_ID = "& numero(codigo_usuario)
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_activa_usuarios
		re.Source = sql
		re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
		re.Open()
		
		if not re.eof and not re.bof then
			email = re("R_EMAIL")
		end if
	
		re.close()
		set re = nothing
	end if
	getEmailUsuario = email
end function

'
' Obtener un valor de un usuario indicando el nombre del dato y el código de usuario
function getDatoUsu(codigo,nombrecorto)
	if ""&typeName(conn_activa_usuarios) = "Connection" then
	
		' Setear grupo
		set miGrupo = setGrupo(getGrupo(session("usuario")))
		Response.Write miGrupo.getAttribute("nombre")
	
		dim re, sql, valor
		valor = ""
		sql = "SELECT R_ID, R_TITULO FROM REGISTROS WHERE R_TITULO = '"& replace(nombre_usuario,"''","'") &"'"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_activa_usuarios
		re.Source = sql
		re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
		re.Open()
		
		if not re.eof and not re.bof then
			valor = re("R_TITULO")
		end if
	
		re.close()
		set re = nothing
	end if
	getDatoUsu = valor
end function

' Obtener el valor de un dato indicado. Para grupos.
function getDatoGrupo(id,dato)
	dim grupo
	set grupo = setGrupo(id)
	if typeOK(grupo) then
'		set nodoGrupo
	end if
end function

'
' Busca si "idGrupo" que le pasamos está en el usuario
function validarGrupo(idGrupo)
	Dim salida, usu, grupo
	salida = false
	for each grupo in usuario.childNodes
		if grupo.nodeName = "grupo" then
			if ""&idGrupo = ""&grupo.getAttribute("id") then
				salida = true
			end if
		end if
	next
	validarGrupo = salida
end function


'
' Obtener permiso para los "Usuario no validados"
function getPermisoNv(zona,idioma)
	' NO DESARROLLADA
	getPermisoNv = TRUE
end function

'
' Comprueba y la sección que le pasmos podemos tiene hijos a los que tenemos permiso
function getPermisoHijos(pCodigo,pLugar)
	getPermisoHijos = true
end function
'
' Funcion que nos devuelve seteado la colección de nodos permiso. Si le pasamos un nombre de permiso nos cogerá sólo los de ese nombre.
function setPermisos(setPermisos_codigo,setPermisos_permiso)
	Dim usu : set usu = setUsuario(setPermisos_codigo)
	if ""&setPermisos_permiso <> "" then
		set setPermisos = usu.selectNodes("permisos/permiso[@nombre='"&setPermisos_permiso&"']")
	else
		set setPermisos = usu.selectNodes("permisos")
	end if
end function

'
' Comprobar la existencia de la zona/cualidad indicada en el idioma indicado, si el idioma es "" comprueba solo la cualidad
' Sabiendo esto podemos desactivas una cualidad (noticias, faqs,...) quitandole los idiomas en su nodo xml
function evalCualidad(evalCualidad_zona,evalCualidad_idioma)

	Dim zona, idioma
	zona = replace(lcase(""&evalCualidad_zona),"acceso_","")
	idioma = lcase(""&evalCualidad_idioma)
	if zona = "edicion" or zona = "nodo_grupos" then
		evalCualidad = true
		exit function
	end if

	dim cualid, esta_cualid
	esta_cualid = false
	for each cualid in cualidades.childNodes
	
		if ""&cualid.nodeName = zona then
			esta_cualid = true
			set cualidad = cualid
		end if
	next
	if esta_cualid then
		if idioma = "" then
			evalCualidad = true
		else
			dim esta_idioma
			esta_idioma = false
			dim idi
			for each idi in cualidad.getElementsByTagName("idioma")
				if ""&idi.getAttribute("nombre") = idioma then
					esta_idioma = true
				end if
			next
			if esta_idioma then
				evalCualidad = true
			else
				evalCualidad = false
			end if			
			
		end if
	else
		evalCualidad = false
	end if
	
end function

'
' obtiene el permiso para el usuario indicado en la cualidad indicada
function getPermisoPara(getPermisoPara_zona, getPermisoPara_idioma, getPermisoPara_codigo)
	getPermisoPara = getPermisoParaRuta(getPermisoPara_zona, getPermisoPara_idioma, getPermisoPara_codigo,"")
end function

'
' obtiene el permiso para el usuario indicado en la cualidad indicada se puede validar en una sección en concreo del XML sección
function getPermisoParaRuta(pZona, pIdioma, pCodigo, pRuta)

	Dim cualidad, zona, idioma, codigo, ruta, usu
	dim re, sql
	zona = lcase(""&pZona)
	idioma = lcase(""&pIdioma)
	codigo = ""&pCodigo
	ruta = ""&pRuta
	ruta = replace(ruta,"\","/")
	ruta = replace(ruta,"//","/")
	if ruta <> "" then
		if inStr(ruta,"/") = 1 then
			ruta = right(ruta,len(ruta)-1)
		end if
		if inStr(len(ruta),ruta,"/") then
			ruta = left(ruta,len(ruta)-1)
		end if
	end if

	' ruta es la sección del XML secciones en la que queremos comprobar el permiso.
	' si es vaicio validaremos solo el permiso de edicion de contenidos global.

	' Lo primero es ver si la cualidad esta creada, activa y con el idioma solicitado
	if not evalCualidad(zona,idioma) then getPermisoParaRuta = false : exit function end if
	' El administrador (id=1) tiene derecho a todo
	if codigo = "1" then getPermisoParaRuta = true : exit function end if

	if ""&typeName(conn_activa_usuarios) = "Connection" then
		sql = "SELECT R_ID, R_SECCION, R_PERMISOS FROM REGISTROS WHERE R_ID = "& numero(pCodigo)
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_activa_usuarios
		re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
		re.Open()
		if re.eof then
			re.close()
			set re = nothing
			getPermisoParaRuta = false
			exit function
		end if
	else
		getPermisoParaRuta = false
		exit function
	end if

	' Si pertnece al grupo (1)Administradores: lo puede todo
	if re("R_SECCION") = 1 then getPermisoParaRuta = true : exit function end if

	' Si es un usuario personalizado  --------------------------------------------------------------------------------------------------------- GRUPO
	if re("R_SECCION") = 2 then
		if ""& re("R_PERMISOS") <> "" then
			set xml_permisos = CreateObject("MSXML.DOMDocument")
			if xml_permisos.LoadXML(re("R_PERMISOS")) then
				set nodo_permisos = xml_permisos.selectSingleNode("permisos")
				if not typeOK(nodo_permisos) then
					getPermisoParaRuta = false
					exit function
				end if
			else
				getPermisoParaRuta = false
				exit function
			end if

			set temp = nodo_permisos.selectSingleNode(zona &"_"& idioma)
			if typeOK(temp) then
				getPermisoParaRuta = true
				exit function
			else
				getPermisoParaRuta = false
				exit function
			end if
		end if
		getPermisoParaRuta = true
		exit function
	end if

	' Si pertenece a un grupo (no personalizado)  --------------------------------------------------------------------------------------------------------- GRUPO
	if re("R_SECCION") >= 3 then
		xq = "//grupos/grupo[@id='"& re("R_SECCION") &"']/permiso[@nombre='"&zona&"']"
		if ruta <> "" then
			xq = xq & "[@ruta='"&ruta&"']"
		end if
		if idioma <> "" then
			xq = xq & "[(@idioma='"&idioma&"' || @idioma='todos')]"
		end if
		'Response.Write vbcrlf & "xq: " & xq
		getPermisoParaRuta = typeOK(xml_grupos.selectNodes(xq).item(0))
	end if

	re.close()
	set re = nothing

end function

'
' Nos devuelve "true" si encuentra el permiso y el idioma que necesita en algun grupo asociado al usuario, devuelve "false" en cualquier otro caso
function getPermiso(getPermiso_zona,getPermiso_idioma)
	dim zona, idioma
	zona = lcase(""&getPermiso_zona)
	idioma = lcase(""&getPermiso_idioma)
	getPermiso = getPermisoPara(zona,idioma,session("usuario"))
end function

'
' Nos devuelve "true" si encuentra el permiso y el idioma que necesita en algun grupo asociado al usuario, devuelve "false" en cualquier otro caso
function listSubpermisos(zona)

	Dim permiso
	Dim migrupo
	set nodoGrupo = setGrupo(session("usuario"))
	if typeOK(nodoGrupo) then
		for each permiso in nodoGrupo.childNodes
			if ""&permiso.getAttribute("nombre") = ""&zona then
				listSubpermisos = permiso.getAttribute("secciones")
				exit function
			end if
		next
	end if
	listSubpermisos = ""
end function

'
' Saber si el usuario que le pasamos tiene derecho de aSkipper sobre los nodo_grupos del grupo que le pasamos
function getPermisoAdminGrupo(codigoUsuario,idGrupo)

	' El grupo "usuario no validados" no se edita
	' El grupo "Admnistrador master" no se edita
	if ""&idGrupo = "0" or ""&idGrupo = "1" then
		getPermisoAdminGrupo = false
		exit function
	end if

	' Marter lo puede todo
	if ""&codigoUsuario = "1" then
		getPermisoAdminGrupo = true
		exit function
	end if

	Dim usu, migrupo
	if codigoUsuario = session("usuario") then
		set usu = usuario
	else
		set usu = setUsuario(codigoUsuario)
		if not typeOK(usu) then
			getPermisoAdminGrupo = false
			exit function
		end if
	end if

	migrupo = ""&usu.getAttribute("grupo")
	Dim permiso, permisos
	if migrupo = "-1" then
		' nodo_grupos personalizados
		set permisos = usu.selectSingleNode("permisos")
		if not typeOK(permisos) then
			getPermisoAdminGrupo = false
			exit function	
		end if
		for each permiso in permisos.childNodes
			if ""&permiso.getAttribute("nombre") = "nodo_grupos" and ""&permiso.getAttribute("grupo") = ""&idGrupo then
				getPermisoAdminGrupo = true
				exit function
			end if
		next
	else
		if typeOK(usu) then
			Dim grupo
			for each grupo in grupos.childNodes
				if grupo.getAttribute("id") = miGrupo then
					for each permiso in grupo.childNodes
						if permiso.nodeName = "permiso" then
							if ""&lcase(permiso.getAttribute("nombre")) = "nodo_grupos" and permiso.getAttribute("grupo") = idGrupo then
								getPermisoAdminGrupo = true
								exit function
							end if
						end if
					next
				end if
			next
		end if
	end if

	getPermisoAdminGrupo = false
end function

'
' Declara el nodo usuaro correspondiente al codigo que le pasamos
function getUsuario(codigo)
	dim usu
	for each usu in nodo_grupos.childNodes
		if ""&usu.getAttribute("codigo") = ""&codigo then
			set getUsuario = usu
			exit function
		end if
	next
	set getUsuario = Nothing
end function

'
' Borrar un usuario a partir del código
function borrarUsuario(codigo)
	' Declaro el nodo
	dim usuBorrar
	set usuBorrar = getUsuario(codigo)
	if not typeOK(usuBorrar) then
		borrarUsuario = false
		exit function
	else
		nodo_grupos.removeChild(usuBorrar)
		if err<>0 then
			borrarUsuario = false
			exit function
		end if
	end if
	borrarUsuario = true
end function

'
' Obtener el nombre de un grupo a partir de su id
function getNombreGrupo(param_id)
	Dim id, grupo : id = ""&param_id
	if id <> "" then
		for each grupo in grupos.childNodes
			if ""&grupo.getAttribute("id") = ""&id then
				getNombreGrupo = grupo.getAttribute("nombre")
				exit function
			end if
		next
	else
		Response.Write "[Vacio]"
	end if
	getNombreGrupo = ""
end function

'
' Setear el nodo idiomas
function setIdiomas()
	Dim nodoIdi
	set nodoIdi = xmlConfig.selectSingleNode("configuracion/idiomas")
	if typeOK(nodoIdi) then
		set setIdiomas = nodoIdi
		exit function
	end if
	set setIdiomas = nothing
end function

'
' Devuelve el nombre completo para las siglas de idioma indicadas (esp,eng, ...)
function getNombreIdioma(idi)
	dim idioma
	idioma = replace(" "&idi," ","")
	if idioma <> "" then
		Dim nodoIdi
		set nodoIdi = xmlConfig.selectSingleNode("configuracion/idiomas/"& idi)
		if typeOK(nodoIdi) then
			getNombreIdioma = ""&nodoIdi.text
		else
			getNombreIdioma = ""
		end if
	else
		getNombreIdioma = ""
	end if
end function








%>