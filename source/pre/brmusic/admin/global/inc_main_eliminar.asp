<%id = request.Form("id")
	if not esNumero(id) then
		unerror = true : msgerror = "No se ha recibido una id."
	else
		sql = "SELECT R_ID, R_FOTO, R_ICONO, R_ARCHIVO, R_SECCION, R_SECCION2"
		sql = sql & " FROM REGISTROS WHERE R_ID = "& id
		
		set conn_activa = server.CreateObject("ADODB.Connection")
		conn_activa.Open conn_
		consultaXOpen sql,2
			if reTotal > 0 then
				foto = ""&re("R_FOTO")
				icono = ""&re("R_ICONO")
				archivo = ""&re("R_ARCHIVO")
				seccion = re("R_SECCION")
				seccion2 = re("R_SECCION2")
				re.delete()
				re.update()
			end if
		consultaXClose()
		
		downRegSeccion(seccion)
		if config_activo_seccion2 then
			downRegSeccion2(seccion2)
		end if
		
		' Compruebo que tiene foto y la borro.
		if foto <> "" then
			foto = server.MapPath("/"& c_s & "datos/" & session("idioma") &"/"& session("cualid")&"/fotos/"&foto)
			if not borrarArchivo(foto) then
				msginfo = msginfo & "<br>Se ha detectado una imagen en el registro de la base de datos pero no se ha logrado borrar.<br> - Puede que no exista."
			end if
		end if
		
		' Compruebo que si tiene icono y lo borro.
		if icono <> "" then
			icono = server.MapPath("/"& c_s & "datos/"& session("idioma") &"/"& session("cualid")&"/iconos/"&icono)
			if not borrarArchivo(icono) then
				msginfo = msginfo & "<br>Se ha detectado un icono en el registro de la base de datos pero no se ha logrado borrar.<br> - Puede que no exista."
			end if
		end if
		
		' Compruebo que si tiene archivo y lo borro.
		if archivo <> "" then
			archivo = server.MapPath("/"& c_s & "datos/" & session("idioma") &"/"& session("cualid")&"/archivos/"&archivo)
			if not borrarArchivo(archivo) then
				msginfo = msginfo & "<br>Se ha detectado un archivo en el registro de la base de datos pero no se ha logrado borrar.<br> - Puede que no exista."
			end if
		end if

		reOrdena()
		mi_seccion = seccion
		reOrdena()
		mi_seccion2 = seccion2
		reOrdena()
		
		' Libero conexión activa
		conn_activa.Close()
		set conn_activa = nothing

		if unerror then
			Response.Write msgerror
		else
			%>
			Un momento ...
			<script>
			try{
				var f = top.frames[1].frames[0].f // Frame de la izquierda
				f.ac.value = ""
				f.action = "main.asp"
				f.target = ""
				f.submit()
				<%if msginfo = "" then%>
				location.href = 'inicio.asp'
				<%end if%>
			}catch(unerror){ alert(unerror.description) }
			</script>
			<br><b>Nota:</b> <%=msginfo%>
			<%
		end if
				
	end if
%>