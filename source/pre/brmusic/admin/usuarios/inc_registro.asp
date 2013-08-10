<!--#include virtual="/admin/inc_sha256.asp" -->
<!--#include virtual="/admin/global/inc_rutinas.asp" -->
<%

	sub campoFormRegistro(usuario)
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

	sub registro_formulario
		' Pintamos un formulario con todos los datos configurados en el XML del grupo para el cual se realizará el registro
		if typeOK(nodoContenido) then
			set nodoFormPlantilla = nodoContenido.selectSingleNode("formularioregistro")
			if not typeOK(nodoFormPlantilla) then
				unerror = true : msgerror = "El nodo ['contenido/formularioregistro'] no se encuentra en el XML de la página actual."
			end if
		end if
		
		if not unerror then
			if numero(admin_grupo) = 0 then
				unerror = true : msgerror = "No se ha definido la ID del grupo al que se debe registrar."
			end if
		end if
		
		' Cargo XML usuario con todos los grupo y seteo el grupo para el registro
		if not unerror then
			ruta_xml_grupos = "/" & c_s & "datos/grupos.xml"
			Set xml_grupos = CreateObject("MSXML.DOMDocument")
			if not xml_grupos.Load(Server.MapPath(ruta_xml_grupos)) then
				unerror = true : msgerror = "No ha sido posible cargar la configuración de los grupos."
			Else
				set nodoGrupo = xml_grupos.selectNodes("/datos/grupos/grupo[@id="& admin_grupo &"]").item(0)
				if not typeOK(nodoGrupo) then
					unerror = true : msgerror = "No se ha encontrado el grupo solicitado en la configuración del registro."
				end if
			End If
		end if

		if not unerror then%>

			<%=getFormulario("",nodoGrupo)%>

		<%end if
		
		if unerror then
			if ""&session("usuario") = "1" then
				Response.Write msgerror
			end if
		end if
	end sub

	dim conn_
	' *******************************************************************************************
	sub sub_registro()
		' Comprobamos que ha escrito su nombre y clave con su correcta repetición.
		nombreUsuario = ""&request.Form("usuario")
		clave = ""&request.Form("clave")
		clave_r = ""&request.Form("clave_r")
		email = ""&request.Form("email")
		
		' Enviar o no un email de confirmación
		confirm_email = cbool(request.Form("confirm_email"))
		' Email del que envia
		dim confirm_email_email : confirm_email_email = ""&request.Form("confirm_email_email")
		' Nombre del que envia
		dim confirm_email_nombre : confirm_email_nombre = ""&request.Form("confirm_email_nombre")
		' Asunto
		dim confirm_email_asunto : confirm_email_asunto = ""&request.Form("confirm_email_asunto")
		' Variable para saber si se ha enviado con éxito
		dim confirm_email_envio
		' Variable del mensaje para el email de confirmación
		dim confirm_email_msg : confirm_email_msg = ""
		
		confirm_email_msg = confirm_email_msg & "Nombre: <b>" & nombreUsuario & "</b><br>"
		confirm_email_msg = confirm_email_msg & "Clave: <b>" & clave & "</b><br>"
		confirm_email_msg = confirm_email_msg & "E-mail: <b>" & email & "</b><br>"
		
		if nombreUsuario = "" then
			registro_error = true : registro_msgerror = "<br>Escriba un nombre de usuario."
		end if
		
		' Comprobar nombre.
		errorValidarNombre = validarNombreUsuario(nombreUsuario)
		if errorValidarNombre <> true then
			registro_error = true : registro_msgerror = "<br>Revise el nombre de usuario:" & errorValidarNombre
		end if
		
		if not registro_error then
			if clave = "" then
				registro_error = true : registro_msgerror = "<br>Escriba una clave de usuario."
			end if
		end if
		
		' Comprobar que la clave es del patron correcto
		errorValidarClave = validarClaveUsuario(clave)
		if errorValidarClave <> true then
			registro_error = true : registro_msgerror = "<br>Revise la clave de usuario:" & errorValidarClave
		end if
		
		' Que las clave coincidan
		if not registro_error then
			if clave <> clave_r then
				registro_error = true : registro_msgerror = "<br>Las claves escritas no coinciden."
			end if
		end if

		' Comprobar que el email es de formato válido
		if not registro_error then
			validaemail = validarEmail(email)
			if validaemail <> true then
				registro_error = true : registro_msgerror = validaemail
			end if
		end if

		' Comprobar que el nombre de usuario no exista (en todos los grupos).
		if not registro_error then
			sql = "SELECT R_TITULO FROM REGISTROS WHERE R_TITULO = '"& nombreUsuario &"'"
			if ""&typeName(conn_activa_usuarios) = "Connection" then
				set re = Server.CreateObject("ADODB.Recordset") : re.ActiveConnection = conn_activa_usuarios : re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
				re.Open()
				if not re.eof then
					registro_error = true : registro_msgerror = "<br>El nombre de usuario escrito está siendo usado por otra persona."
				end if
				re.Close()
				set re = nothing
			end if
		end if

		' Comprobar que no exista el email.
		if not registro_error then
			sql = "SELECT R_EMAIL FROM REGISTROS WHERE R_EMAIL = '"& email &"'"
			if ""&typeName(conn_activa_usuarios) = "Connection" then
				set re = Server.CreateObject("ADODB.Recordset") : re.ActiveConnection = conn_activa_usuarios : re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
				re.Open()
				if not re.eof then
					registro_error = true : registro_msgerror = "<br>El e-mail escrito está siendo usado por otra persona."
				end if
				re.Close()
				set re = nothing
			end if
		end if

		' Recuperamos idGrupo y verificamos que existe, luego declaramos el nodo.
		if not registro_error then
			if numero(admin_grupo) > 0 then
				idGrupo = admin_grupo
			else
				idGrupo = request.Form("grupo")
			end if
			if idGrupo = "" then
				registro_error = true : registro_msgerror = "<br>No se ha recibido el identificador del grupo."			
			end if
		end if
		
		' Declaramos en nodo para el GRUPO elegido
		if not registro_error then
			set grupo = grupos.selectNodes("//datos/grupos/grupo[@id="&idGrupo&"]").item(0)
			if not typeOK(grupo) then
				registro_error = true : registro_msgerror = "<br>No se ha encontrado el grupo indicado."
			end if
		end if

		if not registro_error then
	
			' Compruebo que ha elejido un idioma.
			if admin_idioma <> "" then
				idioma = admin_idioma
			else
				idioma = request.Form("idioma")
			end if
			if idioma = "" then
				registro_error = true : registro_msgerror = registro_msgerror & "<br>Elija un idioma para este usuario."
			end if
			
			fuente = ""
			orden = "999999"
			enportada = 0
			activo = 1
			enlace = ""
			fecha = "0:00:00"
			fechaini = "0:00:00"
			fechafin = "0:00:00"
			piefoto = ""
			pos_foto = ""
			pos_icono = ""

			campo = ""
			sql = "INSERT INTO"

			Dim sql_nombres, sql_valores

			' Campos fijos
			sql_nombres = "R_TITULO, R_CLAVE, R_EMAIL, R_SECCION, R_USUARIO, R_FUENTE, R_ORDEN, R_ORDEN_SECCION, R_ORDEN_SECCION2, R_PORTADA, R_ACTIVO, R_ENLACE, R_FECHA, R_FECHAINI, R_FECHAFIN, R_PIE_FOTO, R_POS_FOTO, R_POS_ICONO"
			sql_valores = "'"& nombreUsuario &"', '"& sha256(clave) &"', '"& email &"', "& admin_grupo &", 0,'"& fuente &"', "& orden &", "& orden &", "& orden &", '"& enportada &"', '"& activo &"', '"& enlace &"', '"& fecha &"', '"& fechaini &"', '"& fechafin &"', '"& piefoto &"', '"& pos_foto &"', '"& pos_icono &"'"
			
			set nodosDatos = grupos.selectNodes("/datos/grupos/grupo[@id="& idGrupo &"]//dato")
			an = 0
			for each dato in nodosDatos
				if dato.nodeName = "dato" then
					campo = UCase(""&dato.getAttribute("campo"))
					valorDato = request.Form(dato.getAttribute("nombrecorto"))
					if campo <> "" then
						an = an + 1
						sql_nombres = sql_nombres & ", R_"& campo &""
						sql_valores = sql_valores & ", '"& replace(valorDato,"'","''") &"'"
						
						' Almacenamos todos los datos en el cuerpo del mensaje para el email de confirmación
						confirm_email_msg = confirm_email_msg & dato.getAttribute("titulo")&": <b>" & valorDato & "</b><br>"
						
						' Si el dato es requerido lo comprobamos
						if dato.getAttribute("requerido") = 1 then
							select case dato.getAttribute("tipo")
							case "email"
								' Validaciones para un campo de E-mail (formato ...)
								validaemail=validarEmail(valorDato)
								if validaemail<>True then
									registro_error = true : registro_msgerror = registro_msgerror &"<br>"& validaemail
								end if
							
							case "texto"
								' Validaciones para un campo de texto
								if valorDato = "" then
									registro_error = true : registro_msgerror = registro_msgerror & "<br>El campo '" & dato.getAttribute("titulo") & "' es requerido."
								end if
	
							case "dni"
								' Validaciones para un campo de E-mail (formato ...)
								if not validarDni(valorDato) then
									if dato.getAttribute("msg") <> "" then
										registro_error = true : registro_msgerror = registro_msgerror & "<br>" & dato.getAttribute("msg")
									else
										registro_error = true : registro_msgerror = registro_msgerror & "<br>El campo <b>" & dato.getAttribute("titulo") & "</b> es requerido."
									end if
								end if
							end select
						end if

					end if
				end if
			next

			if not registro_error  then
				sql = sql & " REGISTROS ("& sql_nombres &")"
				sql = sql & " VALUES ("& sql_valores &")"

				if ""&typeName(conn_activa_usuarios) <> "Connection" then
					registro_error = true : registro_msgerror = registro_msgerror & "<br>La conexión con la base de datos de usuarios no está activa."
				else
					on error resume next
					set oConn = server.CreateObject("ADODB.Connection")
					oConn.Open conn_activa_usuarios
					oConn.execute sql
					oConn.Close
					set oConn = nothing
					if err<>0 then
						registro_error = true : registro_msgerror = registro_msgerror & "<br>Error al ejecutar la sentencia SQL.<br>" & err.description & "<br><b>SQL:</b><br>" & sql
					else

						' Incrementar contador de la sección
						upRegSeccion(admin_grupo)

						conn_ = str_conn_usuarios

						campo_orden = "R_ORDEN_SECCION2"
						reOrdena()
						
						campo_orden = "R_ORDEN_SECCION"
						reOrdena()
						
						campo_orden = "R_ORDEN"
						reOrdena()
					end if
					on error goto 0
				end if
			end if

			if not registro_error then
				' Incluimos la fecha en el email
				confirm_email_msg = confirm_email_msg & "<br>Fecha: <b>" & Date() & "</b><br>"
				' Response.Write "<br><b>:: Email ::</b><br>" & confirm_email_msg
				' Enviar el email
				if confirm_email then
					if confirm_email_email <> "" and confirm_email_nombre <> "" and confirm_email_asunto <> "" then
						on error resume next
						set Mailer = CreateObject("Geocel.Mailer")
						Mailer.AddServer "mail.agrupalia.com", 25
						Mailer.FromAddress = confirm_email_email
						Mailer.FromName = confirm_email_nombre
						Mailer.ContentType = "text/html"
						Mailer.AddRecipient email, nombreUsuario
						Mailer.Subject = confirm_email_asunto
						Mailer.Body = confirm_email_msg
						bSuccess = Mailer.Send()
						on error goto 0
						if bSuccess = False Then
							confirm_email_envio = false
						else
							confirm_email_envio = true
						end if
					end if
				end if
			end if

		end if
		
	end sub
%>