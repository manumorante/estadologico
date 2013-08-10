<%
	sql = "SELECT R_ID,R_FOTO,R_ICONO, R_ARCHIVO, R_TIPOARCHIVO,R_SECCION"
	if config_activo_seccion2 then
		sql = sql & ", R_SECCION2"
	end if
	sql = sql & " FROM REGISTROS WHERE 1=2"
	
	cadena = replace(""&request.Form("comun"),",fin","")
	if inStr(cadena,",")>0 then
		cadena = split(cadena,",")
		for each id in cadena
			if id <> "" then
				sql = sql & " OR R_ID = "& id
			end if
		next
	elseif cadena <> "" and isNumeric(cadena) then
		sql = sql & " OR R_ID = " & cadena
	end if

	consultaXOpen sql,2
	while not re.eof

		foto = ""&re("R_FOTO")
		icono = ""&re("R_ICONO")
		archivo = ""&re("R_ARCHIVO")

		' Compruebo que si tiene foto y la borro.
		if foto <> "" then
			foto = server.MapPath("../"& session("idioma")&"/"&session("cualid")&"/fotos/"&foto)
			call borrarArchivo(foto)
		end if
		
		' Compruebo que si tiene icono y lo borro.
		if icono <> "" then
			icono = server.MapPath("../"& session("idioma")&"/"&session("cualid")&"/iconos/"&icono)
			call borrarArchivo(icono)
		end if
		
		' Compruebo que si tiene archivo y lo borro.
		if archivo <> "" then
			archivo = server.MapPath("../"& session("idioma")&"/"&session("cualid")&"/archivos/"&archivo)
			call borrarArchivo(archivo)
		end if
		
		downRegSeccion(re("R_SECCION"))
		if config_activo_seccion2 then
			downRegSeccion2(numero(re("R_SECCION2")))
		end if


		re.delete
		re.update
		re.movenext
	wend
	consultaXClose()
	
	' Reorganiza el campo R_ORDEN
	reOrdena()

			%>
			Un momento ...
			<script>
			try{
				var f = parent.frames[0].f // Frame de la izquierda
				f.ac.value = ""
				f.action = "main.asp"
				f.target = ""
				f.submit()
				location.href = 'inicio.asp'
			}catch(unerror){
//
			}
			</script>
			