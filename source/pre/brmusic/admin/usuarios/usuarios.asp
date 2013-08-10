<% @LCID = 1034 %>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include file="rutinasParaAdmin.asp" -->
<!--#include virtual="/admin/inc_sha256.asp" -->
<%
archivoXmlusuarios = "/" & c_s & "datos/usuarios.xml"
%>
<html>
<head>
<title>Usuarios</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
</head>
<body class="bodyAdmin">
<%if not unerror then

	if session("usuario") <> "" then
		if not getPermiso("usuarios",session("idioma")) then
			Response.Redirect("noacceso.asp")
		end if
	else
		Response.Redirect("nologeado.asp")
	end if

cadena = ""&request.Form("cadena")
ac = ""&request.Form("ac")
idListGrupo = request.Form("idListGrupo")
%>
<form name="f" action="usuarios.asp" method="post" onSubmit="envio();return false;">
<input type="hidden" name="ac" value="<%=ac%>">
<input type="hidden" name="codigo" value="">
<input type="hidden" name="idListGrupo" value="<%=idListGrupo%>">
<input type="hidden" name="cadenant" value="<%=cadena%>">
<br>
<span class="tituloazonaAdmin">Administraci&oacute;n de Usuarios</span><br>
<br>
Desde aqu&iacute; puede ver y editar los datos de las personas registradas en su web.<br>
Es posible que  tenga acceso s&oacute;lo a determinados grupos de usuarios, siendo
el resto invisibles.<br>
Elija un grupo para ver la lista de usuarios que pertenecen a el. <br>
<br>
<script>
	//
	function lanzapo_win(theURL,winName,ancho,alto,barras) {
		var winl = (screen.width - ancho) / 2;
		var wint = (screen.height - alto) / 2;
		var paramet='top='+wint+',left='+winl+',width='+ancho+',height='+alto+',resizable=yes,scrollbars='+barras+'';
		var splashWin=window.open(theURL,winName,paramet);
		splashWin.focus();
	}

	function nuevoUsuario() {
		f.ac.value = "nuevoUsuario"
		f.cadena.value = ""
		f.submit()
	}
	//
	function envio(){
		f.submit()
	}
	function listGrupo(idGrupo) {
		if (idGrupo != "") {
			f.ac.value = "listaUsuarios"
			f.idListGrupo.value = idGrupo
			f.submit()
		}
	}
	
	function listaCompleta(){
		f.ac.value = "listaUsuarios"
		f.idListGrupo.value = ""
		f.submit()
	}
	function misDatos() {
		f.ac.value = "ampliar"
		f.codigo.value = <%=session("usuario")%>
		f.submit()
	}
	function ir(){
		listGrupo(f.idGrupo.value)
	}
	function buscar(){
		envio()
	}
	function borrarBusqueda(){
		f.cadena.value = ""
		envio()
	}
</script>
<table border="0" cellpadding="2" cellspacing="0">
	<tr>
	<td>
	<select name="idGrupo" class="campoAdmin" onChange="listGrupo(this.value)">
      <option value="">Elija un grupo ...</option>
      <%for each esteGrupo in grupos.childNodes
	  	if getPermisoAdminGrupo(session("usuario"),esteGrupo.getAttribute("id")) then%>
      <option value="<%=esteGrupo.getAttribute("id")%>" <%if request.Form("idListGrupo") = esteGrupo.getAttribute("id") then Response.Write("selected") end if%>><%=esteGrupo.getAttribute("nombre")%></option>
      <%end if
	next%>
    </select></td>
	<td>	  <table border="0" cellpadding="2" cellspacing="0" onClick="ir()" class="botonAdmin" title=" IR ">
      <tr>
        <td>IR</td>
      </tr>
    </table>	  </td>
	<td><table border="0" cellpadding="2" cellspacing="0" onClick="nuevoUsuario()" class="botonAdmin" title=" Crear un nuevo usuario ">
      <tr>
        <td><nobr>Nuevo usuario</nobr></td>
      </tr>
    </table>
	</td>
	<td>	  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table border="0" cellpadding="0" cellspacing="0"  class="botonAdmin" title=" Realizar una búsqueda ">
            <tr>
              <td><nobr><img src="../../spacer.gif" width="1"><input name="cadena" value="<%=cadena%>" type="text" class="campoBuscarUsu" id="cadena" size="30" maxlength="30">
              </nobr></td>
              <td onClick="buscar()"><nobr>&nbsp;&nbsp;Buscar&nbsp;&nbsp;</nobr></td>
            </tr>
          </table></td>

		  <%if cadena <> "" then%>
          <td><table border="0" cellpadding="2" cellspacing="0" onClick="borrarBusqueda()" class="botonAdmin" title=" Borrar búsqueda actual ">
            <tr>
              <td>X</td>
            </tr>
          </table></td>
		  <%end if%>

        </tr>
      </table></td>
	<td><a href="JavaScript:misDatos()" title=" Ver o editar mis datos de usuario "><img src="../global/img/usuario.gif" width="18" height="18" border="0" align="absmiddle"> Mis datos </a></td>
	</tr>
</table>
<br>

<%select case request.Form("ac")
case "nuevoUsuario"%>
<script>
	function chanGrupo(idGrupo) {
		if (idGrupo != "") {
			f.ac.value = "nuevoUsuario"
			f.submit()
		}
	}
	function selectYnuevo(grupo){
		if (grupo != "") {
			f.idListGrupo.value = grupo
			f.ac.value = "nuevoUsuario"
			f.submit()
		}
	}
	function enviarRegistro() {
		f.ac.value = "insertarUsuario"
		f.submit()
	}
	function crearPropiedades(){
		f.ac.value = "insertarUsuario2"
		f.submit()
	}
	//
	function popSeccionesEdicion(){
		lanzapo_win("permisos_edicion.asp","SeccionesEdicion",250,500,1)
	}
</script>
<%
	c_idGrupo = request.Form("grupo")
	idListGrupo = ""&request.Form("idListGrupo")

	if idListGrupo <> "" then
		c_idGrupo = idListGrupo
	end if
	

	' Usuario personalizado (Ningún grupo)
	if idListGrupo = "-1" then%>

		<span class="tituloazonaAdmin">Nuevo Usuario Personalizado</span>
		<br>
		<br>
		Escoja las cualidades que desea para este usuario.<br>
		Los campos marcados con uns asterisco (*) son obligatorios.<br>
		<br>
		<table  border="0" cellspacing="0" cellpadding="1">
		  <tr bordercolor="#849ACE" class="fondoOscuroAdmin">
		    <td height="30" colspan="3"><span class="colorBlanco">&nbsp;&nbsp;Propiedades del usuario </span></td>
	      </tr>
		  <tr>
		    <td colspan="3">&nbsp;</td>
	      </tr>
		  <tr class="campoAdmin">
            <td>&nbsp;Cualidad&nbsp;</td>
            <td>&nbsp;Idiomas/Grupos</td>
            <td>Configurar</td>
		  </tr>

		<%for each cualid in cualidades.childNodes
		
			if cualid.nodeName <> "idiomas" and cualid.nodeName <> "infogeneral"then
				if claseFila = "clasefila1" then
					claseFila = "clasefila2"
				else
					claseFila = "clasefila1"
				end if%>
				
				<tr class="<%=claseFila%>">
				<td>&nbsp;<%=cualid.getAttribute("nombre")%></td>
				<td>
				<%
				select case cualid.nodeName
				' USUARIOS
				case "usuarios"
					for each grupo in grupos.childNodes
						if grupo.getAttribute("id") <> 1 and grupo.getAttribute("id") <> 0 then%>
							<nobr><input name="usuarios_<%=grupo.getAttribute("id")%>" id="usuarios_<%=grupo.getAttribute("id")%>" type="checkbox" value="1"><label for="usuarios_<%=grupo.getAttribute("id")%>"><%=grupo.getAttribute("nombre")%></label></nobr>
						<%end if
					next

				' EDICIÓN
				case "edicion"
					
				' RESTO DE CUALIDADES
				case else
					for each idi in cualid.getElementsByTagName("idioma")%>
						<input id="idioma_<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>" name="idioma_<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>" type="checkbox" value="1"><label for="idioma_<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>"><%=getNombreIdioma(idi.getAttribute("nombre"))%></label>
					<%next
				end select%>&nbsp;</td>

				<td>
				<%
				select case cualid.nodeName
				' USUARIOS
				case "usuarios"

				' EDICIÓN
				case "edicion"%>
					<input name="secciones_esp" type="hidden" id="secciones_esp" value="">
					<input type="button" class="botonAdmin" onClick="popSeccionesEdicion()" value="Secciones">
				<%' RESTO DE CUALIDADES
				case else
				end select%>
				</td>
				</tr>
			<%end if
		next%>
		<tr><td colspan="4" align="right">&nbsp;</td></tr>
		<tr class="fondoOscuroAdmin">
		  <td colspan="4" align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Volver">
		  <input name="" type="button" class="botonAdmin" onClick="crearPropiedades()" value="Enviar"></td>
		  </tr>
		</table>
		
		<br>

	<%elseif idListGrupo <> "" then
		' Declaro el NODO GRUPO
		for each g in grupos.childNodes
			if c_idGrupo = g.getAttribute("id") then
				set grupo = g
			end if
		next%>
		<input type="hidden" name="grupo" value="<%=c_idGrupo%>">

		<span class="tituloazonaAdmin">Nuevo Usuario</span>
		<br><br>
		Introduzca los datos necesarios para crear un nuevo usuario.<br>
		Los campos marcados con uns asterisco (*) son obligatorios.<br>
		<br>
		
		<%
		'Si tenemos un grupo seleccionado
		if typeName(grupo) <> "Nothing" and typeName(grupo) <> "Empty" then%>
		<table border="0" cellpadding="1" cellspacing="0" class="fondoOscuroAdmin">
			<tr>
			<td height="30"><font color="#FFFFFF">&nbsp;GRUPO: <b><%=getNombreGrupo(c_idGrupo)%></b></font></td>
			</tr>
			<tr>
			<td>
			<table border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
				<tr>
				<td>
				<table border="0" align="center" cellpadding="2" cellspacing="0"><tr bgcolor="#F7F9FB">
				  <td align="right"><b>Idioma prinicpal*</b>: </td>
				  <td>
				<%
				dim idiomas
				set idiomas = setIdiomas()
				for each idi in idiomas.childNodes
				%>
					<input type="radio" name="idioma" id="id_<%=idi.nodeName%>" value="<%=idi.nodeName%>"><label for="id_<%=idi.nodeName%>"><%=idi.text%></label>
				<%next%></td>
				  <tr>
					<td align="right"><nobr><b>Nombre de usuario</b>*:</nobr></td>
					<td><input name="usuario" type="text" class="campoAdmin" id="c_usuario" value="" size="15" maxlength="200"></td>
				  <tr>
					<td align="right"><b>Clave</b>*: </td>
					<td><input name="clave" type="password" class="campoAdmin" id="c_clave" value="" size="10" maxlength="50"></td>
				  </tr>
					<tr>
					<td align="right"><nobr><b>Repetir clave</b>*:</nobr> </td>
					<td><input name="claver" type="password" class="campoAdmin" id="c_clave" value="" size="10" maxlength="50"></td>
					</tr>
					<tr>
					<td align="right"><b>E-mail</b>*: </td>
					<td><input name="email" type="text" class="campoAdmin" id="c_clave" value="" size="30" maxlength="100"></td>
					</tr>
					<%
					for each dato in grupo.childNodes
						if dato.nodeName = "dato" then%>
							<tr>
							<td align="right"><%if dato.getAttribute("requerido") = 1 then%><b><%=dato.getAttribute("nombre")%></b>*<%else%><b><%=dato.getAttribute("nombre")%></b><%end if%>:</td>
							<td><% campoForm dato%></td>
							</tr>
						<%elseif dato.nodeName = "bloque" then
							if ""&dato.getAttribute("titulo") <> "" then%>
							<tr>
							<td align="right"><font color="#4564AD" size="2"><%=dato.getAttribute("titulo")%></font></td>
							<td>&nbsp;</td>
							</tr>
							<%end if%>
							<tr>
							<td colspan="2" align="right"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td bgcolor="#849ACE"><img src="../../spacer.gif" width="1" height="1"></td>
                              </tr>
                            </table></td>
							</tr>
						<%end if
					next%></table></td>
			  </tr>
			  </table></td>
		  </tr>
					<tr>
					<td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Volver">
					<input type="button" class="botonAdmin" onClick="enviarRegistro()" value="Enviar"></td>
					</tr></table>
				<%else%>
					<b>Debe elejir el grupo en el que desea crear el nuevo usuario</b>.
				<%end if
			end if
	

case "cambiarPermisos2"
	codigo = request.Form("codigo")
	set usu = setUsuario(codigo)
	if not typeOK(usu) then
	else
		' Borramos el nodo con  los permisos acuales
		set permisos = usu.selectSingleNode("permisos")
		usu.removeChild(permisos)
	
		' Cualidades e idioma seleccionados
		set nodoPermisos = xmlObj.createElement("permisos")
		for each a in request.Form()
			if inStr(a,"idioma_") then
				corte = split(a,"_")
				set nodoPermiso = xmlObj.createElement("permiso")
				' Cualidad
				set att = xmlObj.createAttribute("nombre")
				att.nodeValue = lcase(corte(1))
				nodoPermiso.setAttributeNode(att)
				set att = nothing
				' Idioma
				set att = xmlObj.createAttribute("idioma")
				att.nodeValue = lcase(corte(2))
				nodoPermiso.setAttributeNode(att)
				set att = nothing

				nodoPermisos.appendChild(nodoPermiso)
				set nodoPermiso = nothing
			elseif inStr(a,"usuarios_") then
				corte = split(a,"_")
				set nodoPermiso = xmlObj.createElement("permiso")
				' Cualidad
				set att = xmlObj.createAttribute("nombre")
				att.nodeValue = "usuarios"
				nodoPermiso.setAttributeNode(att)
				set att = nothing
				' Grupo
				set att = xmlObj.createAttribute("grupo")
				att.nodeValue = lcase(corte(1))
				nodoPermiso.setAttributeNode(att)
				set att = nothing
				nodoPermisos.appendChild(nodoPermiso)
				set nodoPermiso = nothing
			end if
		next

		' Secciones de edición
		arrSecciones = split(request.Form("secciones_esp"),"|")
		for each a in arrSecciones
			if ""&a <> "" then
				set nodoPermiso = xmlObj.createElement("permiso")
	
				' Cualidad
				set att = xmlObj.createAttribute("nombre")
				att.nodeValue = "edicion"
				nodoPermiso.setAttributeNode(att)
				set att = nothing
				
				' Idioma
				set att = xmlObj.createAttribute("idioma")
				att.nodeValue = "esp"
				nodoPermiso.setAttributeNode(att)
				set att = nothing
	
				' Ruta
				set att = xmlObj.createAttribute("ruta")
				att.nodeValue = a
				nodoPermiso.setAttributeNode(att)
				set att = nothing
	
				nodoPermisos.appendChild(nodoPermiso)
				set nodoPermiso = nothing
			end if
		next
		
		usu.appendChild(nodoPermisos)
		xmlObj.save Server.MapPath(archivoXmlusuarios)
	end if
	%>
	<script>
	f.ac.value = "preeditar"
	f.codigo.value = <%=codigo%>
	f.submit()</script>
	<%
 
case "cambiarPermisos"
codigo = ""&request.Form("codigo")

set usuEdi = setUsuario(codigo)

%>
<br>
<script language="javascript" type="text/javascript">
	function editarPropiedades(){
		f.ac.value = "cambiarPermisos2"
		f.codigo.value = <%=codigo%>
		f.submit()
	}
	//
	function popSeccionesEdicion(){
		lanzapo_win("permisos_edicion.asp?P="+f.secciones_esp.value,"SeccionesEdicion",250,500,1)
	}
</script>

                <table  border="0" cellspacing="0" cellpadding="1">
                  <tr bordercolor="#849ACE" class="fondoOscuroAdmin">
                    <td height="30" colspan="3"><span class="colorBlanco">&nbsp;&nbsp;Editar
                        propiedades
                        del usuario </span></td>
                  </tr>
                  <tr>
                    <td colspan="3">&nbsp;</td>
                  </tr>
                  <tr class="campoAdmin">
                    <td>&nbsp;Cualidad&nbsp;</td>
                    <td>&nbsp;Idiomas</td>
                    <td>&nbsp;Configurar</td>
                  </tr>
			<%for each cualid in cualidades.childNodes
			
			if cualid.nodeName <> "idiomas" and cualid.nodeName <> "infogeneral" then
				if claseFila = "clasefila1" then
					claseFila = "clasefila2"
				else
					claseFila = "clasefila1"
				end if%>
                  <tr class="<%=claseFila%>">
                    <td>&nbsp;<%=cualid.getAttribute("nombre")%></td>
                    <td>
				<%
				
				select case cualid.nodeName
				case "usuarios"
				
					for each grupo in grupos.childNodes
						if grupo.getAttribute("id") <> 1 and grupo.getAttribute("id") <> 0 then%>
							<nobr><input name="usuarios_<%=grupo.getAttribute("id")%>" id="usuarios_<%=grupo.getAttribute("id")%>" type="checkbox" value="1" <% if getPermisoAdminGrupo(codigo,grupo.getAttribute("id")) then Response.Write "checked" end if %>><label for="usuarios_<%=grupo.getAttribute("id")%>"><%=grupo.getAttribute("nombre")%></label></nobr>
						<%end if
					next
				case "edicion"
					Response.Write ""
				case else
					for each idi in cualid.getElementsByTagName("idioma")%>
						<input name="idioma_<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>" id="idioma_<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>" type="checkbox" value="1" <% if getPermisoPara(cualid.nodeName,idi.getAttribute("nombre"),codigo) then Response.Write "checked" end if %>><label for="idioma_<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>"><%=getNombreIdioma(idi.getAttribute("nombre"))%></label>
					<%next
				end select
				%>&nbsp;</td>
                    <td>
				<%
				select case cualid.nodeName
				' USUARIOS
				case "usuarios"

				' EDICIÓN
				case "edicion"%>
				<%
				
				set permisos = setPermisos(codigo,"edicion")
				str_permisos = "|"
				for each a in permisos
					str_permisos = str_permisos  & a.getAttribute("ruta") & "|"
				next%>
					<input name="secciones_esp" type="hidden" id="secciones_esp" value="<%=str_permisos%>">
					<input type="button" class="botonAdmin" onClick="popSeccionesEdicion('<%=str_permisos%>')" value="Secciones">
				<%' RESTO DE CUALIDADES
				case else
				end select%>
					</td>
                  </tr>
                  <%end if
		next%>
                  <tr>
                    <td colspan="4" align="right">&nbsp;</td>
                  </tr>
                  <tr class="fondoOscuroAdmin">
                    <td colspan="4" align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Volver">
                    <input name="" type="button" class="botonAdmin" onClick="editarPropiedades()" value="Enviar"></td>
                  </tr>
                </table>
<%

case "insertarUsuario3"
	' ...::: INSERTAR USUARIO PERSONALIZADO :::...
	
	' Comprobamos que ha escrito su nombre y clave con su correcta repetición.
		nombreUsuario = ""&request.Form("usuario")
		clave = ""&request.Form("clave")
		claver = ""&request.Form("claver")
		email = ""&request.Form("email")
		idgrupo = "-1"
		idioma = ""&request.Form("idioma")
		
		if idioma = "" then
			unerror = true : msgerror = "<br>Indique un <b>idioma principal</b>."
		end if

		if nombreUsuario = "" then
			unerror = true : msgerror = "<br>Escriba un <b>nombre</b> de usuario."
		end if
		
		' Comprobar que el nombre no tiene espacios.
		errorValidarNombre = validarNombreUsuario(nombreUsuario)
		if errorValidarNombre <> true then
			unerror = true : msgerror = "<br><b>Revise el nombre de usuario</b>:" & errorValidarNombre
		end if

		' Comprobar que el nombre de usuario no exista. Ni siquiera en otro para otro grupo (pues podria cambiar de grupo y estaria repetido entonces).
		if not unerror then
			for each otro in usuarios.childNodes
				if lcase(""&otro.getAttribute("usuario")) = lcase(""&nombreUsuario) then
					miNombreYaExiste = true
				end if
			next
			if miNombreYaExiste then
				unerror = true : msgerror = "<br>El <b>nombre de usuario</b> escrito está siendo <b>usado</b> por otra persona."
			end if
		end if
		
		if not unerror then
			if clave = "" then
				unerror = true : msgerror = "<br>Escriba una <b>clave</b> de usuario."
			end if
		end if
		
		' Comprobar que la clave es del patron correcto
		errorValidarClave = validarClaveUsuario(clave)
		if errorValidarClave <> true then
			unerror = true : msgerror = "<br><b>Revise la clave de usuario</b>: " & errorValidarClave
		end if
		
		' Que las clave coincidan
		if not unerror then
			if clave <> claver then
				unerror = true : msgerror = "<br>Las claves escritas <b>no coinciden</b>."
			end if
		end if
		
		' Comprobar que el email es de formato válido
		if not unerror then
			validaemail = validarEmail(email)
			if validaemail <> true then
				unerror = true : msgerror = validaemail
			end if
		end if
		
		' Que el email no exista por otro usaurio
		if not unerror then
			for each otro in usuarios.childNodes
				if lcase(""&otro.getAttribute("email")) = lcase(""&email) then
					miEmailYaExiste = true
				end if
			next
			if miEmailYaExiste then
				unerror = true : msgerror = "<br>El <b>e-mail</b> escrito está siendo <b>usado</b> por otra persona."
			end if
		end if
	
		if not unerror then
			set grupo = setGrupo("-1")
			if not typeOK(grupo) then
				unerror = true : msgerror = "No se ha encontrado el nodo con los datos para los usuarios personalizados"
			end if
		end if

		if not unerror then
			set nuevoUsuario = xmlObj.createElement("usuario")
			for each dato in grupo.childNodes
				if dato.nodeName = "dato" then

					valorDato = request.Form(dato.getAttribute("nombrecorto"))
					' Si el dato es requerido lo comprobamos
					if dato.getAttribute("requerido") = 1 then
						select case dato.getAttribute("tipo")
						case "email"
							' Validaciones para un campo de E-mail (formato ...)
							validaemail=validarEmail(valorDato)
							if validaemail<>True then
								formerror = true : msgform = msgform & validaemail
							end if

						case "texto"
							' Validaciones para un campo de texto
							if valorDato = "" then
								formerror = true : msgform = msgform & "<br>El campo <b>" & dato.getAttribute("nombre") & "</b> es requerido."
							end if

						case "dni"
							' Validaciones para un campo de E-mail (formato ...)
							if not validarDni(valorDato) then
								if dato.getAttribute("msg") <> "" then
									formerror = true : msgform = msgform & "<br>" & dato.getAttribute("msg")
								else
									formerror = true : msgform = msgform & "<br>El campo <b>" & dato.getAttribute("nombre") & "</b> es requerido."
								end if
							end if
						end select
					end if
					
					' Creamos los datos
					set nuevoDato = xmlObj.createElement("dato")
					set attNombre = xmlObj.createAttribute("nombre")
					set attNombreCorto = xmlObj.createAttribute("nombrecorto")
					nuevoDato.setAttributeNode(attNombre)
					nuevoDato.setAttributeNode(attNombreCorto)
					attNombre.nodeValue = dato.getAttribute("nombre")
					attNombreCorto.nodeValue = dato.getAttribute("nombrecorto")
					if ""&valorDato <> "" then
						nuevoDato.text = valorDato
					end if				
					nuevoUsuario.appendChild(nuevoDato)
					set attNombreCorto = nothing
					set attNombre = nothing
					set nuevoDato = nothing
				end if
			next
			
			' Detemos el proceso si en la anterior rutina faltan campos requeridos
			'  y pintamos el informe que nos ha declaro en tal caso.
			if formerror then
				unerror = true : msgerror = msgform
			else

				' Cualidades e idioma seleccionados
				set nodoPermisos = xmlObj.createElement("permisos")
				
				for each a in request.Form()
					if inStr(a,"idioma_") then
						corte = split(a,"_")
						set nodoPermiso = xmlObj.createElement("permiso")
						' Cualidad
						set att = xmlObj.createAttribute("nombre")
						att.nodeValue = lcase(corte(1))
						nodoPermiso.setAttributeNode(att)
						set att = nothing
						' Idioma
						set att = xmlObj.createAttribute("idioma")
						att.nodeValue = lcase(corte(2))
						nodoPermiso.setAttributeNode(att)
						set att = nothing

						nodoPermisos.appendChild(nodoPermiso)
						set nodoPermiso = nothing
					elseif inStr(a,"usuarios_") then
						corte = split(a,"_")
						set nodoPermiso = xmlObj.createElement("permiso")
						' Cualidad
						set att = xmlObj.createAttribute("nombre")
						att.nodeValue = "usuarios"
						nodoPermiso.setAttributeNode(att)
						set att = nothing
						' Idioma
						set att = xmlObj.createAttribute("idioma")
						att.nodeValue = "todos"
						nodoPermiso.setAttributeNode(att)
						set att = nothing
						' Grupo
						set att = xmlObj.createAttribute("grupo")
						att.nodeValue = lcase(corte(1))
						nodoPermiso.setAttributeNode(att)
						set att = nothing
						nodoPermisos.appendChild(nodoPermiso)
						set nodoPermiso = nothing
					elseif inStr(a,"edicion_") then
						corte = split(a,"_")
						set nodoPermiso = xmlObj.createElement("permiso")
						' Cualidad
						set att = xmlObj.createAttribute("nombre")
						att.nodeValue = "edicion"
						nodoPermiso.setAttributeNode(att)
						set att = nothing
						
						' Idioma
						set att = xmlObj.createAttribute("idioma")
						att.nodeValue = lcase(corte(1))
						nodoPermiso.setAttributeNode(att)
						set att = nothing
						
						' Sección
						set att = xmlObj.createAttribute("ruta")
						att.nodeValue = request.Form(a)
						nodoPermiso.setAttributeNode(att)
						set att = nothing

						nodoPermisos.appendChild(nodoPermiso)
						set nodoPermiso = nothing
					end if
				next
				nuevoUsuario.appendChild(nodoPermisos)
				
				
				' Localizo y declaro el ATRIBUTO MAXCODIGO al que llamaré "attMaxcodigo".
				for each att in usuarios.attributes
					if att.nodeName = "maxcodigo" then
						set attMaxcodigo = att
					end if
				next

				' Compruebo que esta y si no está lo creo.
				if typeName(attMaxcodigo) = "Nothing" or typeName(attMaxcodigo) = "Empty" then
					set attMaxcodigo = xmlObj.createAttribute("maxcodigo")
					attMaxcodigo.nodeValue = 0
					usuarios.setAttributeNode(attMaxcodigo)
				end if
				
				' Incremento el ATTIBUTO MAXCODIGO para la cuenta de usuarios
				attMaxcodigo.nodeValue = attMaxcodigo.nodeValue + 1
				
				' Le metemos el ATTIBUTO FECHA
				set attFecha = xmlObj.createAttribute("fecharegistro")
				attFecha.nodeValue = Date()
				nuevoUsuario.setAttributeNode(attFecha)
				
				' Le metemos el ATTIBUTO IDIOMA
				set attIdioma = xmlObj.createAttribute("idioma")
				attIdioma.nodeValue = idioma
				nuevoUsuario.setAttributeNode(attIdioma)
				set attIdioma = nothing

				' Le metemos el ATTIBUTO CODIGO que le corresponde según incremento (almacenado en nodo "usuarios" del XML principal).
				set attCodigo = xmlObj.createAttribute("codigo")
				maxcodigo = cint(usuarios.getAttribute("maxcodigo"))
				attCodigo.nodeValue = attMaxcodigo.nodeValue
				nuevoUsuario.setAttributeNode(attCodigo)
				set attCodigo = nothing

				' Le metemos el ATTIBUTO GRUPO
				set attGrupo = xmlObj.createAttribute("grupo")
				attGrupo.nodeValue = idgrupo
				nuevoUsuario.setAttributeNode(attGrupo)
				set attGrupo = nothing

				' Meto ATRIBUTO USUARIO
				set attUsuario = xmlObj.createAttribute("usuario")
				attUsuario.nodeValue = nombreUsuario
				nuevoUsuario.setAttributeNode(attUsuario)
				set attUsuario = nothing
				
				' Meto ATRIBUTO CLAVE
				set attClave = xmlObj.createAttribute("clave")
				attClave.nodeValue = SHA256(clave)
				nuevoUsuario.setAttributeNode(attClave)
				set attClave = nothing
				
				' Meto ATRIBUTO EMAIL
				set attEmail = xmlObj.createAttribute("email")
				attEmail.nodeValue = email
				nuevoUsuario.setAttributeNode(attEmail)
				set attEmail = nothing

				' Meto mi nuevo "nodo usuario" en el nodo "usuario" del XML principal
				usuarios.appendChild(nuevoUsuario)
			
				' Guardo y aviso de posibles fallos		
				on error resume next
					xmlObj.save Server.MapPath(archivoXmlusuarios)
					if err <> 0 then
						unerror = true : msgerror = "Se ha producido un error al intentar guardar en el archivo XML.<br>"&err.Description
					end if
				on error goto 0
			end if


		end if
		
		
		
		if unerror then%>
		
<table border="0" cellpadding="1" cellspacing="0" bgcolor="#990000">
          <tr>
            <td bgcolor="#FFFFFF"><font color="#990000">Por favor,
                revise los siguientes datos:</font></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
                <tr>
                  <td bgcolor="#FEFAFA"><%=msgerror%><br><br></td>
                </tr>
            </table>            </td>
          </tr>
          <tr>
            <td align="right" bgcolor="#F5F5F5"><input type="button" class="botonAdmin" onClick="history.back()" value="Volver"></td>
          </tr>
  </table>
			
		<%else%>
		<b>Usuario insertado</b>.<br>
		Cargando lista completa ...
		<script>
			f.ac.value = "listaUsuarios"
			f.idListGrupo.value = "-1"
			f.submit()
		</script>
		<%end if

case "insertarUsuario2"%>
	<script>
		function registrar(){
			f.ac.value = "insertarUsuario3"
			f.submit()
		}
	</script>
  Ya tenemos toda la imformaci&oacute;n sobre las cualidades<br>
		        <br>

	<%for each a in request.Form()
		if inStr(a,"idioma_") or inStr(a,"usuarios_") then%>
			*<input name="<%=a%>" type="hidden" value="<%=request.Form(a)%>"><%=vbCr%>
		<%end if
	next
	
	secciones_esp = ""&request.Form("secciones_esp")
	if inStr(secciones_esp,"|") then
		arrSeccionesEsp = Split(secciones_esp,"|")
		n=0
		for each a in arrSeccionesEsp
			if ""&a <> "" then
				n=n+1%>
				<input name="edicion_esp_<%=n%>" type="hidden" value="<%=a%>">
			<%end if
		next
	end if


	
	set grupo = setGrupo("-1")
	if not typeOK(grupo) then
		unerror = true : msgerror = "No se ha encontrado los datos predefinidos de los usuarios personalizados."
	end if

	if not unerror then%>

	<table border="0" cellpadding="1" cellspacing="0" class="fondoOscuroAdmin">
      <tr>
        <td height="30"><b><font color="#FFFFFF">&nbsp;&nbsp;Datos
        del usuario personalizado</font></b></td>
      </tr>
      <tr>
        <td>
          <table  border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
            <tr>
              <td><table border="0" cellpadding="2" cellspacing="0">
                <tr bgcolor="#F7F9FB">
                  <td align="right"><b>Idioma prinicpal*</b>: </td>
                  <td>
                    <%
					  set idiomas = setIdiomas()
					  for each idi in idiomas.childNodes%>
                    <input type="radio" name="idioma" id="id_<%=idi.nodeName%>" value="<%=idi.nodeName%>"><label for="id_<%=idi.nodeName%>"><%=idi.text%></label>
                    <%next%>
                  </td>
                  <%
					for each dato in grupo.childNodes
						if dato.nodeName = "dato" then%>
                  <%end if
					next%>
                <tr>
                  <td align="right"><nobr><b>Nombre de usuario</b>*:</nobr></td>
                  <td><input name="usuario" type="text" class="campoAdmin" id="c_usuario" value="" size="15" maxlength="200"></td>
                <tr>
                  <td align="right"><b>Clave</b>*: </td>
                  <td><input name="clave" type="password" class="campoAdmin" id="c_clave" value="" size="10" maxlength="50"></td>
                </tr>
                <tr>
                  <td align="right"><nobr><b>Repetir clave</b>*:</nobr> </td>
                  <td><input name="claver" type="password" class="campoAdmin" id="c_clave" value="" size="10" maxlength="50"></td>
                </tr>
                <tr>
                  <td align="right"><b>E-mail</b>*: </td>
                  <td><input name="email" type="text" class="campoAdmin" id="c_clave" value="" size="30" maxlength="100"></td>
                </tr>
                  <%for each dato in grupo.childNodes
		if dato.nodeName = "dato" then%>
                  <tr>
                    <td align="right">
							<%if dato.getAttribute("requerido") = 1 then%>
								<b><%=dato.getAttribute("nombre")%></b>*
							<%else%>
								<b><%=dato.getAttribute("nombre")%></b>
							<%end if%>:</td>
                    <td><% campoForm dato%></td>
                  </tr>
                  <%end if
	next%>
              </table></td>
            </tr>
          </table></td>
      </tr>
      <tr>
        <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Volver">
        <input name="" type="button" class="botonAdmin" onClick="registrar()" value="Enviar"></td>
      </tr>
  </table>
	<%end if
	
case "insertarUsuario"

		' Comprobamos que ha escrito su nombre y clave con su correcta repetición.
		nombreUsuario = ""&request.Form("usuario")
		clave = ""&request.Form("clave")
		claver = ""&request.Form("claver")
		email = ""&request.Form("email")
		idgrupo = ""&request.Form("grupo")

		if clave = "" then
			unerror = true : msgerror = "<br>Escriba una <b>clave</b> de usuario."
		else
			' Comprobar que la clave es del patron correcto
			errorValidarClave = validarClaveUsuario(clave)
			if errorValidarClave <> true then
				unerror = true : msgerror = "<br><b>Revise la clave de usuario</b>: " & errorValidarClave
			else
				' Que las clave coincidan
				if clave <> claver then
					unerror = true : msgerror = "<br>Las claves escritas <b>no coinciden</b>."
				end if
			end if
		end if
	
		if nombreUsuario = "" then
			unerror = true : msgerror = "<br>Escriba un <b>nombre</b> de usuario."
		end if
		
		' Comprobar que el nombre no tiene espacios.
		errorValidarNombre = validarNombreUsuario(nombreUsuario)
		if errorValidarNombre <> true then
			unerror = true : msgerror = "<br><b>Revise el nombre de usuario</b>:" & errorValidarNombre
		end if
		
		' Comprobar que el nombre de usuario no exista. Ni siquiera en otro para otro grupo (pues podria cambiar de grupo y estaria repetido entonces).
		if not unerror then
			for each otro in usuarios.childNodes
				if lcase(""&otro.getAttribute("usuario")) = lcase(""&nombreUsuario) then
					miNombreYaExiste = true
				end if
			next
			if miNombreYaExiste then
				unerror = true : msgerror = "<br>El <b>nombre de usuario</b> escrito está siendo <b>usado</b> por otra persona."
			end if
		end if

		' Compruebo que ha elejido un idioma.
		idioma = request.Form("idioma")
		if idioma = "" then
			unerror = true : msgerror = "<br>Elija un <b>idioma</b> para este usuario."
		end if
		
		' Comprobar que el email es de formato válido
		if not unerror then
			validaemail = validarEmail(email)
			if validaemail <> true then
				unerror = true : msgerror = validaemail
			end if
		end if
		
		' Que el email no exista por otro usaurio
		if not unerror then
			for each otro in usuarios.childNodes
				if lcase(""&otro.getAttribute("email")) = lcase(""&email) then
					miEmailYaExiste = true
				end if
			next
			if miEmailYaExiste then
				unerror = true : msgerror = "<br>El <b>e-mail</b> escrito está siendo <b>usado</b> por otra persona."
			end if
		end if

		' Verificamos el grupo, luego declaramos el nodo.
		if not unerror then
			if idGrupo = "" then
				unerror = true : msgerror = "<br>No se ha recibido el <b>identificador</b> del grupo."			
			end if
		end if
		
		' Declaramos en nodo para el GRUPO elegido
		if not unerror then
			for each este in grupos.childNodes
				if ""&este.getAttribute("id") = idGrupo then
					set grupo = este				
				end if
			next
			' Comprobamos que se ha encontrado y declarado correctamente el GRUPO indicado.
			if not typeOK(grupo) then
				unerror = true : msgerror = "<br>No se ha encontrado el grupo indicado."
			end if
		end if

		' Creamos el NUEVO NODO y lo llenamos de los datos que nos llegan del formulario
		if not unerror then
			set nuevoUsuario = xmlObj.createElement("usuario")
			
			for each dato in grupo.childNodes
				if dato.nodeName = "dato" then

					valorDato = request.Form(dato.getAttribute("nombrecorto"))
					' Si el dato es requerido lo comprobamos
					if dato.getAttribute("requerido") = 1 then
						select case dato.getAttribute("tipo")
						case "email"
							' Validaciones para un campo de E-mail (formato ...)
							validaemail=validarEmail(valorDato)
							if validaemail<>True then
								
									formerror = true : msgform = msgform & validaemail
								
							end if

						case "texto"
							' Validaciones para un campo de texto
							if valorDato = "" then
								formerror = true : msgform = msgform & "<br>El campo <b>" & dato.getAttribute("nombre") & "</b> es requerido."
							end if

						case "dni"
							' Validaciones para un campo de E-mail (formato ...)
							if not validarDni(valorDato) then
								if dato.getAttribute("msg") <> "" then
									formerror = true : msgform = msgform & "<br>" & dato.getAttribute("msg")
								else
									formerror = true : msgform = msgform & "<br>El campo <b>" & dato.getAttribute("nombre") & "</b> es requerido."
								end if
							end if
						end select
					end if
					
					' Creamos los datos
					set nuevoDato = xmlObj.createElement("dato")
					set attNombre = xmlObj.createAttribute("nombre")
					set attNombreCorto = xmlObj.createAttribute("nombrecorto")
					nuevoDato.setAttributeNode(attNombre)
					nuevoDato.setAttributeNode(attNombreCorto)
					attNombre.nodeValue = dato.getAttribute("nombre")
					attNombreCorto.nodeValue = dato.getAttribute("nombrecorto")
					if valorDato <> "" then
						nuevoDato.text = valorDato
					end if				
					nuevoUsuario.appendChild(nuevoDato)
					set attNombre = nothing
					set nuevoDato = nothing
				end if
			next
			
			' Detemos el proceso si en la anterior rutina faltan campos requeridos
			'  y pintamos el informe que nos ha declaro en tal caso.
			if formerror then
				unerror = true : msgerror = msgform
			else

				' Localizo y declaro el ATRIBUTO MAXCODIGO al que llamaré "attMaxcodigo".
				for each att in usuarios.attributes
					if att.nodeName = "maxcodigo" then
						set attMaxcodigo = att
					end if
				next

				' Compruebo que esta y si no está lo creo.
				if typeName(attMaxcodigo) = "Nothing" or typeName(attMaxcodigo) = "Empty" then
					set attMaxcodigo = xmlObj.createAttribute("maxcodigo")
					attMaxcodigo.nodeValue = 0
					usuarios.setAttributeNode(attMaxcodigo)
				end if
				
				' Incremento el ATTIBUTO MAXCODIGO para la cuenta de usuarios
				attMaxcodigo.nodeValue = attMaxcodigo.nodeValue + 1
				
				' Le metemos el ATTIBUTO FECHA
				set attFecha = xmlObj.createAttribute("fecharegistro")
				attFecha.nodeValue = Date()
				nuevoUsuario.setAttributeNode(attFecha)
				set attFecha = nothing

				' Le metemos el ATTIBUTO IDIOMA
				set attIdioma = xmlObj.createAttribute("idioma")
				attIdioma.nodeValue = idioma
				nuevoUsuario.setAttributeNode(attIdioma)
				set attIdioma = nothing

				' Le metemos el ATTIBUTO CODIGO que le corresponde según incremento (almacenado en nodo "usuarios" del XML principal).
				set attCodigo = xmlObj.createAttribute("codigo")
				maxcodigo = cint(usuarios.getAttribute("maxcodigo"))
				attCodigo.nodeValue = attMaxcodigo.nodeValue
				nuevoUsuario.setAttributeNode(attCodigo)
				set attCodigo = nothing

				' Le metemos el ATTIBUTO GRUPO
				set attGrupo = xmlObj.createAttribute("grupo")
				attGrupo.nodeValue = idgrupo
				nuevoUsuario.setAttributeNode(attGrupo)
				set attGrupo = nothing

				' Meto ATRIBUTO USUARIO
				set attUsuario = xmlObj.createAttribute("usuario")
				attUsuario.nodeValue = nombreUsuario
				nuevoUsuario.setAttributeNode(attUsuario)
				set attUsuario = nothing
				
				' Meto ATRIBUTO CLAVE Codificada en SHA256
				set attClave = xmlObj.createAttribute("clave")
				attClave.nodeValue = SHA256(clave)
				nuevoUsuario.setAttributeNode(attClave)
				set attClave = nothing
				
				' Meto ATRIBUTO EMAIL
				set attEmail = xmlObj.createAttribute("email")
				attEmail.nodeValue = email
				nuevoUsuario.setAttributeNode(attEmail)
				set attEmail = nothing

				' Meto mi nuevo "nodo usuario" en el nodo "usuario" del XML principal
				usuarios.appendChild(nuevoUsuario)
			
				' Guardo y aviso de posibles fallos		
				on error resume next
					xmlObj.save Server.MapPath(archivoXmlusuarios)
					if err <> 0 then
						unerror = true : msgerror = "Se ha producido un error al intentar guardar en el archivo XML.<br>"&err.Description
					end if
				on error goto 0
			end if
		
		end if


		if unerror then%>
		<table border="0" cellpadding="1" cellspacing="0" bgcolor="#990000">
          <tr>
            <td bgcolor="#FFFFFF"><font color="#990000">Por favor,
                revise los siguientes datos:</font></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
                <tr>
                  <td bgcolor="#FEFAFA"><%=msgerror%><br><br></td>
                </tr>
            </table>            </td>
          </tr>
          <tr>
            <td align="right" bgcolor="#F5F5F5"><input type="button" class="botonAdmin" onClick="history.back()" value="Volver"></td>
          </tr>
        </table>
		<br>
  <%else
			Response.Redirect("usuarios.asp?msg=Usuario introducido correctamente")
		end if%>

<%case "listaUsuarios"%>
<script>
	function borrarUsuario(codigo,nombre){
		if(confirm("¿Seguro que desea borrar el usuario \""+ nombre +"\"?")){
			f.ac.value = "borrar"
			f.codigo.value = codigo
			f.submit()
		}
	}

	function editarUsuario(codigo){
		f.ac.value = "preeditar"
		f.codigo.value = codigo
		f.submit()
	}
	
	function cambiarPass(codigo){
		f.ac.value = "cambiarpass"
		f.codigo.value = codigo
		f.submit()
	}

	function ampliar(codigo) {
		f.ac.value = "ampliar"
		f.codigo.value = codigo
		f.submit()
	}
</script>
<span class="tituloazonaAdmin">Lista de usuarios</span><br>
<br>

<table width="600" border="0" cellpadding="0" cellspacing="1">
<tr class="campoAdmin">
  <td>Nombre</td>
  <td>Grupo</td>
  <td width="100%" align="right">Opciones</td>
</tr>
	<%idGrupo = ""&request.Form("idListGrupo")
	' Número de coincidencias
	numCoin = 0
	for each esteUsu in usuarios.childNodes

		pertenece = true
		' Ningún grupom, o sea, usuario personalizado
		if ""&esteUsu.getAttribute("grupo") <> idGrupo then
			pertenece = false
		end if
		
		' Si se ha realizado una búsqueda
		if cadena <> "" then
			if inStr(""&lcase(esteUsu.getAttribute("usuario")),cadena) <= 0 then
				pertenece = false
			end if
		end if

		' Si pertenece al grupo que hemos solicitado listas
		if pertenece then
		
			if claseFila = "clasefila1" then
				claseFila = "clasefila2"
			else
				claseFila = "clasefila1"
			end if
		
			numCoin = numCoin + 1
			
		%>
		

              <tr class="<%=claseFila%>">
                <td align="left" bgcolor="#FFFFFF">&nbsp;&nbsp;<a href="#" onClick="ampliar(<%=esteUsu.getAttribute("codigo")%>)"><%=esteUsu.getAttribute("usuario")%></a>&nbsp;</td>
                <td><nobr>&nbsp;<%=getNombreGrupo(esteUsu.getAttribute("grupo"))%></nobr></td>
                <td align="right"><table border="0" cellspacing="0" cellpadding="1">
                    <tr>
                      <td><input name="" type="button" class="botonAdmin" title=" Cambiar clave " onClick="cambiarPass(<%=esteUsu.getAttribute("codigo")%>)" value="PASS"></td>
                      <td><a href="JavaScript:editarUsuario(<%=esteUsu.getAttribute("codigo")%>)"><img src="../global/img/lapiz.gif" alt=" Editar " width="18" height="18" border="0"></a></td>
                      <td><%if esteUsu.getAttribute("codigo") = 1 or esteUsu.getAttribute("codigo") = session("usuario")  then%><img src="../global/img/papelera.gif" alt=" No se puede borrar " width="18" height="18"><%else%><a href="JavaScript:borrarUsuario(<%=esteUsu.getAttribute("codigo")%>,'<%=esteUsu.getAttribute("usuario")%>')"><img src="../global/img/papelera.gif" alt=" Borrar " width="18" height="18" border="0"></a><%end if%></td>
                    </tr>
                </table></td>
              </tr>
  <%end if
	next%>
	
	<tr>
	<td colspan="3" class="fondoAdmin"><table width="100%"  border="0" cellspacing="0" cellpadding="4">
      <tr>
        <td>Encontrado(s) <b><%=numCoin%></b> resultado(s)
          <%if cadena <> "" then%>
que contenga(n) <b><%=cadena%></b>
<%end if%>
.</td>
      </tr>
    </table>	  </td>
	</tr>
  </table>

		



<%
case "borrar"

	' una funcioncita que borre ...
	codigo = request.Form("codigo")
	if codigo <> "" then
		if borrarUsuario(codigo) then
			on error resume next
				xmlObj.save Server.MapPath(archivoXmlusuarios)
				if err=0 then%>
					<script>
					f.ac.value = "listaUsuarios"
					f.submit()
					</script>
				<%else
					Response.Redirect("usuarios.asp?msg=No se ha podido guardar en el xml.")
				end if
			on error goto 0
		else
			Response.Redirect("usuarios.asp?msg=Error, no se pudo borrar el usuario indicado.")
		end if
	else
		Response.Redirect("usuarios.asp?msg=No se ha recibido un código de usuario.")
	end if
		%>
	
<%case "ampliar"%>
<script>
	function editar(codigo) {
		f.ac.value = "preeditar"
		f.codigo.value = codigo
		f.submit()
	}
	
	function permisos(codigo) {
		f.ac.value = "cambiarPermisos"
		f.codigo.value = codigo
		f.submit()
	}
</script>
<%
	' Código
	codigo = request.Form("codigo")

	if codigo = "" then unerror = true : msgerror = "No se ha recibido ningún código." end if
	' Usuario
	if not unerror then
		set usuAmpli = getUsuario(codigo)
		if typeName(usuAmpli) = "Nothing" then unerror = true : msgerror = "No se ha encontrado el código de usuario indicado." end if
	end if

	if not unerror then%>
	
		<span class="tituloazonaAdmin">Información extendida</span><br><br>
		<table border="0" cellpadding="1" cellspacing="0" class="fondoOscuroAdmin">
          <tr>
            <td height="30"><b><font color="#FFFFFF">&nbsp;&nbsp;Datos del usuario</font></b></td>
          </tr>
          <tr>
            <td>
              <table  border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
                <tr>
                  <td><table border="0" cellpadding="2" cellspacing="0">
<tr>
  <td align="right"><nobr><b>Idioma principal</b>:</nobr></td>
  <td><%=getNombreIdioma(usuAmpli.getAttribute("idioma"))%></td>
  <td align="right">id:<%=usuAmpli.getAttribute("codigo")%></td>
  <%
					idgrupo = getGrupo(codigo)
					%>
  <%
'set grupo = setGrupo(idgrupo)
'Response.Write typeName(grupo)
'for each dato in grupo.childNodes
'	if dato.nodeName = "dato" then
'		Response.Write dato.getAttribute("nombre") & "<br>"
'	end if
'next

for each dato in usuAmpli.childNodes
	if dato.nodeName = "dato" then%>
  <%end if
next%>
<tr>
                  <td align="right"><nobr><b>Nombre de usuario</b>:</nobr></td>
                  <td colspan="2"><%=usuAmpli.getAttribute("usuario")%></td>
                      
                    <tr>
                        <td align="right"><b>E-mail</b>: </td>
                        <td colspan="2"><%=usuAmpli.getAttribute("email")%></td>
                    </tr>
					
					<%
					idgrupo = getGrupo(codigo)
					%>
					
                      <tr>
                        <td align="right"><b>Grupo</b>:</td>
                        <td colspan="2"><%=getNombreGrupo(idgrupo)%>&nbsp;</td>
                      </tr>
					  
					  <%if ""&idgrupo = "-1" then%>
				     <tr>
                        <td align="right" valign="top"><b>Permisos</b>:</td>
                        <td colspan="2" valign="top">
						<%
						set permisos = usuAmpli.selectSingleNode("permisos")
						if typeOK(permisos) then%>
						<table width="100%"  border="0" cellpadding="0" cellspacing="0" class="fondoOscuroAdmin">
                          <tr>
                            <td><table width="100%" border="0" cellpadding="0" cellspacing="1">
                              <tr bgcolor="#FFFFFF">
                                <td align="center" bgcolor="#FFFFFF" class=fondoOscuroCblancoAdmin><b>Permiso</b></td>
                                <td align="center" bgcolor="#FFFFFF" class=fondoOscuroCblancoAdmin><b>Grupo</b>/<b>Idioma/Ruta</b></td>
                              </tr>
                              <%for each permiso in permisos.childNodes%>
                              <tr bgcolor="#FFFFFF">
                                <td>&nbsp;<%=getTituloCualidad(permiso.getAttribute("nombre"))%>&nbsp;</td>
                                <td>&nbsp;
								<%if ""&permiso.getAttribute("idioma") <> "" then%>
								<font color="#777777">Idioma:</font> <%=getNombreIdioma(permiso.getAttribute("idioma"))%>
								<%end if%>
								<%if ""&permiso.getAttribute("grupo") <> "" then%>
								<font color="#777777">Grupo:</font> <%=getNombreGrupo(permiso.getAttribute("grupo"))%>
								<%end if%>
								<%if ""&permiso.getAttribute("ruta") <> "" then%>
								<font color="#777777">Ruta:</font> <%=permiso.getAttribute("ruta")%>
								<%end if%>&nbsp;</td>
                              </tr>
                              <%next%>
                            </table></td>
                          </tr>
                        </table>
						<%end if
						%></td>
					 </tr>
					 <%end if%>
					 
<%

for each dato in usuAmpli.childNodes
	if dato.nodeName = "dato" then%>
		<tr><td align="right"><b><%=dato.getAttribute("nombre")%></b>:</td>
		<td colspan="2"><%=dato.text%></td></tr>
		<%elseif dato.nodeName = "bloque" then
			if ""&dato.getAttribute("titulo") <> "" then%>
			<tr>
			<td align="right"><font color="#4564AD" size="2"><%=dato.getAttribute("titulo")%></font></td>
			<td>&nbsp;</td>
			</tr>
			<%end if%>
			<tr>
			<td colspan="2" align="right"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td bgcolor="#849ACE"><img src="../../spacer.gif" width="1" height="1"></td>
			  </tr>
			</table></td>
			</tr>
		<%end if
next%>




                    
                  </table></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Volver">
              <%if idgrupo = "-1" then%>
			  <input name="Botón" type="button" class="botonAdmin" onClick="permisos(<%=usuAmpli.getAttribute("codigo")%>)" value="Permisos">
              <%end if%>
			  <input type="button" class="botonAdmin" onClick="editar(<%=usuAmpli.getAttribute("codigo")%>)" value="Editar"></td>
          </tr>
        </table>

	<%end if%>



<%case "preeditar"
	' Código
	codigo = request.Form("codigo")
	if codigo = "" then unerror = true : msgerror = "No se ha recibido ningún código." end if
	' Usuario
	if not unerror then
		set usuEdi = getUsuario(codigo)
		if typeName(usuEdi) = "Nothing" then unerror = true : msgerror = "No se ha encontrado el código de usuario indicado." end if
	end if

	' Grupo
	if not unerror then
		idgrupo = getGrupo(codigo)
		set grupo = setGrupo(idgrupo)
		if typeOK(grupo) then
			nombreGrupo = grupo.getAttribute("nombre")
		else
			nombreGrupo = "Ningún grupo"
		end if
	end if	
	
	if unerror then
		Response.Write("<b>Error:</b><br>" & msgerror)
	else
%>
		
	Datos requeridos para el grupo <b><%=nombreGrupo%><br>
	<br>
	</b>
	<script>
		function enviar() {
			f.ac.value = "editar"
			f.codigo.value = <%=codigo%>
			f.submit()
		}
		
		function cambiarClave() {
			f.ac.value = "cambiarpass"
			f.codigo.value = <%=codigo%>
			f.submit()
		}
		
		function cambiarPermisos() {
			f.ac.value = "cambiarPermisos"
			f.codigo.value = <%=codigo%>
			f.submit()
		}
		function volver(){
			f.ac.value = "listaUsuarios"
			f.submit()
		}
		
	</script>
	<table border="0" cellpadding="1" cellspacing="0" class="fondoOscuroAdmin">
      <tr>
        <td height="30"><font color="#FFFFFF">&nbsp;&nbsp;<b>Editar usuario </b></font></td>
      </tr>
      <tr>
        <td><table border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
            <tr>
              <td><table border="0" cellpadding="2" cellspacing="0">
                <tr>
                  <td align="right"><nobr><b>Idioma principal</b>:&nbsp;</nobr></td>
                  <td>
				  <%
				  set idiomas = setIdiomas()
				  n=0
				  for each idi in idiomas.childNodes
				  	n=n+1%>
					  <input type="radio" name="idioma" id="i<%=n%>" value="<%=idi.nodeName%>" <%if idi.nodeName = usuEdi.getAttribute("idioma") then Response.Write "checked" end if%>><label for="i<%=n%>"><%=idi.text%></label>
				  <%next%></td>
                </tr>
                <tr>
                  <td align="right"><nobr><b>Nombre de usuario</b>:&nbsp;</nobr></td>
                  <td><%=usuEdi.getAttribute("usuario")%>
                  <input name="grupo" type="hidden" id="grupo" value="<%=idgrupo%>"></td>
                </tr>
                <tr>
                  <td align="right"><nobr><b>Clave</b>:&nbsp;</nobr></td>
                  <td>
                    <input type="button" class="botonAdmin" onClick="cambiarClave()" value="Editar"></td>
                </tr>
				
				<%if ""&usuEdi.getAttribute("grupo") = "-1" then%>
				<tr>
                  <td align="right"><nobr><b>Permisos</b>:&nbsp;</nobr></td>
                  <td>
                    <input name="" type="button" class="botonAdmin" onClick="cambiarPermisos()" value="Editar"></td>
			    </tr>
				<%end if%>
				<tr>
                  <td align="right"><nobr><b>E-mail</b>*:&nbsp;</nobr></td>
                  <td><input name="email" type="text" class="campoAdmin" id="email" value="<%=usuEdi.getAttribute("email")%>" size="30" maxlength="100"></td>
                </tr>
                <%
				' [[[ esto aun no esta hecho ]]]
				if typeOK(grupo) then
				for each dato in grupo.childNodes
			if dato.nodeName = "dato" then%>
                
                <tr>
                  <td align="right"><nobr><b><%=dato.getAttribute("nombre")%></b><%if dato.getAttribute("requerido") = 1 then%>*<%end if%>:&nbsp;</nobr></td>
                  <td><%campoFormLleno dato,codigo%></td>
                </tr>
			<%elseif dato.nodeName = "bloque" then
				if ""&dato.getAttribute("titulo") <> "" then%>
				<tr>
				<td align="right"><font color="#4564AD" size="2"><%=dato.getAttribute("titulo")%></font></td>
				<td>&nbsp;</td>
				</tr>
				<%end if%>
				<tr>
				<td colspan="2" align="right"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td bgcolor="#849ACE"><img src="../../spacer.gif" width="1" height="1"></td>
				  </tr>
				</table></td>
				</tr>
			<%end if

		next
		end if
		%>
              </table></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td align="right"><input name="" type="button" class="botonAdmin" onClick="volver()" value="Volver">
        <input type="button" class="botonAdmin" onClick="enviar()" value="Enviar"></td>
      </tr>
    </table>
	<%end if
	
case "editar"

	' Usuario
	codigo = request.Form("codigo")
	email = ""&request.Form("email")
	idgrupo = ""&request.Form("grupo")
	idioma = ""&request.Form("idioma")
	
	if idioma = "" then
		unerror = true : msgerror = "<br>Indique un <b>idioma principal</b>."
	end if
	
	' formato de Email
	validaemail = validarEmail(email)
	if validaemail<>True then
		unerror = true : msgerror = validaemail
	end if
	
	' Comprobar q no se esté usando el mismo email
	if not unerror then
		for each otro in usuarios.childNodes
			if lcase(""&otro.getAttribute("email")) = lcase(""&email) and ""&codigo <> ""&otro.getAttribute("codigo") then
				miEmailYaExiste = true
			end if
		next
		if miEmailYaExiste then
			unerror = true : msgerror = "<br>El <b>e-mail</b> escrito está siendo <b>usado</b> por otra persona."
		end if
	end if

	set usuEdi = getUsuario(codigo)
	if typeName(usuEdi) = "Nothing" then unerror = true : msgerror = "No se ha encontrado el usuario indicado." end if

	' Email
	set att = xmlObj.createAttribute("email")
	usuEdi.setAttributeNode(att)
	att.nodevalue = email
	set att = nothing
	
	' Idioma
	set att = xmlObj.createAttribute("idioma")
	usuEdi.setAttributeNode(att)
	att.nodevalue = idioma
	set att = nothing
	
	' última modificacion
	set att = xmlObj.createAttribute("modificado")
	usuEdi.setAttributeNode(att)
	att.nodevalue = date()
	set att = nothing

	if not unerror then
		set grupo = setGrupo(idGrupo)
		if not typeOK(grupo) then
			unerror = true : msgerror = "No se ha encontrado el nodo para el ID GRUPO indicado."
		end if
	end if


	if not unerror then
		
		for each dato in grupo.childNodes
			if dato.nodeName = "dato" then
				valorDato = request.Form(dato.getAttribute("nombrecorto"))
				' Si el dato es requerido lo comprobamos
				if dato.getAttribute("requerido") = 1 then
					select case dato.getAttribute("tipo")
					case "email"
						' Validaciones para un campo de E-mail (formato ...)
							validaemail=validarEmail(valorDato)
							if validaemail<>True then
								
									formerror = true : msgform = msgform & validaemail
								
							end if

					case "texto"
						' Validaciones para un campo de texto
						if valorDato = "" then
							formerror = true : msgform = msgform & "<br>El campo <b>" & dato.getAttribute("nombre") & "</b> es requerido."
						end if

					case "dni"
						' Validaciones para un campo de E-mail (formato ...)
						if not validarDni(valorDato) then
							if dato.getAttribute("msg") <> "" then
								formerror = true : msgform = msgform & "<br>" & dato.getAttribute("msg")
							else
								formerror = true : msgform = msgform & "<br>El campo <b>" & dato.getAttribute("nombre") & "</b> es requerido."
							end if
						end if
					end select
				end if ' requerido
				
				nombre = dato.getAttribute("nombre")
				' Procedo XML:
				



			encontrado = False
			for each nodo in usuEdi.childNodes
				if nodo.nodeName = "dato" then
					if lcase(""&nodo.getAttribute("nombre")) = lcase(""&nombre) then
						nodo.text = valorDato
						encontrado = True
					end if
				end if
			next
			if encontrado=False then
				set nodonuevo = xmlObj.createElement("dato")
				usuEdi.appendChild(nodonuevo)
				nodonuevo.text=valorDato
				
				
				Set atributonuevo= xmlObj.createAttribute("nombre")
				nodonuevo.setAttributeNode(atributonuevo)
				atributonuevo.nodevalue=lcase(""&nombre)
				
			end if
		end if
		next
	end if
	
	if formerror then
		unerror = true : msgerror = msgform
	else
		' Guardo y aviso de posibles fallos		
		on error resume next
			xmlObj.save Server.MapPath(archivoXmlusuarios)
			if err <> 0 then
				unerror = true : msgerror = "Se ha producido un error al intentar guardar en el archivo XML.<br>"&err.Description
			end if
		on error goto 0
	end if
	
	if unerror then%>
		<table border="0" cellpadding="1" cellspacing="0" bgcolor="#990000">
          <tr>
            <td bgcolor="#FFFFFF"><font color="#990000">Por favor,
                revise los siguientes datos:</font></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
                <tr>
                  <td bgcolor="#FEFAFA"><%=msgerror%><br>
                      <br></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td align="right" bgcolor="#F5F5F5"><input type="button" class="botonAdmin" onClick="history.back()" value="Volver"></td>
          </tr>
        </table>
		<%
	else
		Response.Redirect("usuarios.asp?msg=Usuario modificado correctamente")
	end if

case "cambiarpass"
	' Código
	codigo = request.Form("codigo")
	if codigo = "" then unerror = true : msgerror = "No se ha recibido ningún código." end if
	' Usuario
	if not unerror then
		set usuEdi = getUsuario(codigo)
		if typeName(usuEdi) = "Nothing" then unerror = true : msgerror = "No se ha encontrado el código de usuario indicado." end if
	end if
	if unerror then
		Response.Write msgerror
	else
	%>
	<script>
		function enviar() {
			if (f.nuevaclave.value == ""){
				alert("Escriba un clave.")
				f.nuevaclave.focus()
			} else if (f.nuevaclave.value != f.nuevaclave_r.value) {
				alert("Las clave escritas no coinciden.")
			} else {
				f.ac.value = "cambiarpass2"
				f.codigo.value = <%=codigo%>
				f.submit()
			}
		}
	</script>
<br>
<table border="0" cellpadding="1" cellspacing="0" class="fondoOscuroAdmin">
  <tr>
    <td height="30"><font color="#FFFFFF">&nbsp;&nbsp;<b>Cambiar clave</b> </font></td>
  </tr>
  <tr>
    <td><table border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
          <td><table border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td colspan="2">Escriba la nueva <b>clave</b> para el usuario <b><%=usuEdi.getAttribute("usuario")%></b>.</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td align="right"><nobr><b>Clave</b>:&nbsp;</nobr></td>
              <td><input name="nuevaclave" type="password" class="campoAdmin" id="nuevaclave" size="15" maxlength="20">              </td>
            </tr>
            <tr>
              <td align="right"><nobr><b>Repetir clave</b>:&nbsp;</nobr></td>
              <td><input name="nuevaclave_r" type="password" class="campoAdmin" id="nuevaclave_r" size="15" maxlength="20">
              </td>
            </tr>
          </table>            </td>
        </tr>
    </table>
</td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Volver">
    <input type="button" class="botonAdmin" onClick="enviar()" value="Cambiar"></td>
  </tr>
</table>
<br>
&nbsp;
  <%end if

case "cambiarpass2" ' --------------------------------------------------------------------------------------------

	' Código
	codigo = request.Form("codigo")
	if codigo = "" then unerror = true : msgerror = "No se ha recibido ningún código." end if
	' Usuario
	if not unerror then
		set usuEdi = getUsuario(codigo)
		if typeName(usuEdi) = "Nothing" then unerror = true : msgerror = "No se ha encontrado el código de usuario indicado." end if
	end if
	
	if unerror then
		Response.Write msgerror
	else
		set miClave = usuEdi.attributes.getNamedItem("clave")
		miClave.nodeValue = SHA256(request.Form("nuevaclave"))

		' Guardo y aviso de posibles fallos		
		on error resume next
			xmlObj.save Server.MapPath(archivoXmlusuarios)
			if err <> 0 then
				unerror = true : msgerror = "Se ha producido un error al intentar guardar en el archivo XML.<br>"&err.Description
			else
				Response.Redirect("usuarios.asp?msg=Clave cambiada correctamente.")
			end if
		on error goto 0
	end if

case else
	if request.QueryString("msg") <> "" then%>
	<br>
	<table border="0" cellspacing="0" cellpadding="4">
      <tr>
        <td bgcolor="#FFFFFF" class="fondoOscuroAdmin">&nbsp;</td>
        <td bgcolor="#FFFFFF"><font color="#006600"><%=request.QueryString("msg")%></font></td>
      </tr>
  </table>
	<br>
	<%end if
end select

%>
</form>

<%else%>
	<b>Ha ocurrido un error</b><br>
	<%=msgerror%>
<%end if%>

<!--#'include file="rutinasParaAdmin_fin.asp" -->

</body>
</html>
