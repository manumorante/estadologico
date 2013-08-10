<%
	' DEPENDIENTE DE ...
	' *** /inc_rutinas.asp
	' ***  conn_ Cadena de conexion a la BD de la cualidad deseada.

	' Insertar un nuevo registro
	function insertarRegistro(titulo, seccion, seccion2, usuario, fuente, alfinal, enportada, activo, enlace, fecha, fechaini, fechafin, resto_nombres, resto_valores, conn)

			Dim sql_nombres, sql_valores
			sql = "INSERT INTO"

			if alfinal then
				orden = "9999"
			else
				orden = "0.9"
			end if
			
			if ""&fecha = "" then
				fecha = "0:00:00"
			end if

			if ""&fechaini = "" then
				fechaini = "0:00:00"
			end if

			if ""&fechafin = "" then
				fechafin = "0:00:00"
			end if

			' Campos fijos
			sql_nombres = "R_TITULO, R_SECCION, R_SECCION2, R_USUARIO, R_FUENTE, R_ORDEN, R_PORTADA, R_ACTIVO, R_ENLACE, R_FECHA, R_FECHAINI, R_FECHAFIN"
			sql_valores = "'"& titulo &"', "& seccion &", "& seccion2 &", "& usuario &",'"& fuente &"', "& orden &", "& enportada &", "& activo &", '"& enlace &"', '"& fecha &"', '"& fechaini &"', '"& fechafin &"'"

			sql_nombres = sql_nombres & resto_nombres
			sql_valores = sql_valores & resto_valores
			
			sql = sql & " REGISTROS ("& sql_nombres &")"
			sql = sql & " VALUES ("& sql_valores &")"
			conn_ = conn
			insertar = exeSql(sql, conn)
			if ""& insertar = "" then
				upRegSeccion(seccion)
				upRegSeccion2(seccion2)
				reOrdena()
			end if
			insertarRegistro = insertar

	end function

	
	' Insertar un nuevo registro y devulvo su id
	function insertarSeccion(nombre, alfinal, alias, bloqueada, renombrar, eliminar)

			Dim sql_nombres, sql_valores
			sql = "INSERT INTO"

			if alfinal then
				orden = "9999"
			else
				orden = "0.9"
			end if
			
			if ""&fecha = "" then
				fecha = "0:00:00"
			end if

			if ""&fechaini = "" then
				fechaini = "0:00:00"
			end if

			if ""&fechafin = "" then
				fechafin = "0:00:00"
			end if

			' Campos fijos
			sql_nombres = "S_NOMBRE, S_BLOQUEADA, S_RENOMBRAR, S_ELIMINAR, S_ALIAS, S_ORDEN"
			sql_valores = "'"& nombre &"', "& bloqueada &", "& renombrar &","& eliminar &", '"& alias &"', "& orden

			sql = sql & " SECCIONES ("& sql_nombres &")"
			sql = sql & " VALUES ("& sql_valores &")"

			call exeSql(sql, conn_)
			reOrdenaSecciones()
			
			consultaXOpen "SELECT S_ID FROM SECCIONES ORDER BY S_ID DESC",1
			insertarSeccion = re("S_ID")
			consultaXClose()

	end function

	sub consultaUsuarios(sql,tipoBloqueo)
		ruTotal = 0
		if ""&sql = "" then
			unerror = true : msgerror = "La consulta SQL está vacia."
			exit sub
		end if
		set ru = Server.CreateObject("ADODB.Recordset")
		on error resume next

		if ""&typeName(conn_activa_usuarios) = "Connection" then
			ru.ActiveConnection = conn_activa_usuarios
		else
			unerror = true : msgerror = "No hay una conexión activa o la cadena de conexíon esta vacia o es incorrecta."
			exit sub
		end if

		if err<>0 then
			unerror = true : msgerror = "No se ha encontrado la base de datos."
			exit sub
		end if

		ru.Source = sql : ru.CursorType = 1 : ru.CursorLocation = 2 : ru.LockType = tipoBloqueo
		ru.Open()
		if err<>0 then
			unerror = true : msgerror = "La consulta SQL ha devuelto un error.<br><b>SQL:</b>" & sql
			exit sub
		end if
		ruTotal = ru.recordcount
		on error goto 0
	end sub

	sub consultaXOpen(sql,tipoBloqueo)
		reTotal = 0
		if ""&sql = "" then
			unerror = true : msgerror = "La consulta SQL está vacia."
			exit sub
		end if

		set re = Server.CreateObject("ADODB.Recordset")
		'on error resume next

		if ""&typeName(conn_activa) = "Connection" then
			re.ActiveConnection = conn_activa
		elseif ""& conn_ <> "" then
			re.ActiveConnection = conn_
		else
			unerror = true : msgerror = "No hay una conexión activa o la cadena de conexíon esta vacia o es incorrecta."
			exit sub
		end if

		if err<>0 then
			unerror = true : msgerror = "No se ha encontrado la base de datos.CONN: " & conn_
			exit sub
		end if

		re.Source = sql : re.CursorType = 1 : re.CursorLocation = 2 : re.LockType = tipoBloqueo
		re.Open()
		if err<>0 then
			unerror = true : msgerror = "La consulta SQL ha devuelto un error.<br><b>SQL:</b>" & sql
			exit sub
		end if

		reTotal = re.recordcount
		on error goto 0
	end sub
	
	sub consultaXClose()
		on error resume next
			re.Close() : set re = Nothing
		on error goto 0
	end sub

	sub consultaUsuariosClose()
		on error resume next
			ru.Close() : set ru = Nothing
		on error goto 0
	end sub
	
	' Ordena los campos R_orden
	sub reOrdena()
		dim n, re, reTotal, sql
		dim misecc
		dim misecc2

		misecc = numero(mi_seccion)
		misecc2 = numero(mi_seccion2)
		
		' Campo orden
		if misecc>0 and misecc2>0 then
			campo_orden = "R_ORDEN_SECCION2"
			where = "WHERE R_SECCION = "& misecc &" AND R_SECCION2 = "& misecc2
		elseif misecc>0 then
			campo_orden = "R_ORDEN_SECCION"
			where = "WHERE R_SECCION = "& misecc
		else
			campo_orden = "R_ORDEN"
		end if
		sql = "SELECT "& campo_orden &" FROM REGISTROS "& where &" ORDER BY "& campo_orden
		
		reTotal = 0
		on error resume next
		set re = Server.CreateObject("ADODB.Recordset")
		if ""&typeName(conn_activa) = "Connection" then
			re.ActiveConnection = conn_activa
		else
			re.ActiveConnection = conn_
		end if

		if err<>0 then
			unerror = true : msgerror = "Error en conexión a base de datos.<br>No se ha podido reordenar.<br>"&err.description&"<br>"&conn_
		else
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			reTotal = re.recordcount
			if err<>0 then
				unerror = true : msgerror = "Sql:<br>" & sql &"<br><br>Error:<br>"&err.description
			end if
		end if
		on error goto 0
		
		if reTotal > 0 then
			n = 1
			while not re.eof
				re(campo_orden) = n
				n = n + 1
				re.Update()
				re.MoveNext()
			wend
		end if
		set re = nothing
	end sub
	
	' Generación de XML configurados
	sub genXml()
	
		if not typeOK(nodoCualid) then
			Response.Write "No se ha detectado una cualidad."
		else
			dim mfso, mfsoescribir
			d = chr(34) ' comillas dobles
			if not unerror then
				dim mixml
				for each x in nodoCualid.selectNodes("xml")
					'Response.Write "<br>2) Localizado nodo cualidad."
					mixml = "<?xml version="&d&"1.0"&d&" encoding="&d&"iso-8859-1"&d&"?>" & vbCrLf
					dim secci : set secci = Server.CreateObject("ADODB.Recordset")
					secci.ActiveConnection = conn_
					dim subsecci : set subsecci = Server.CreateObject("ADODB.Recordset")
					subsecci.ActiveConnection = conn_
					mixml = mixml & "<datos>" & vbCrLf
					sql = ""&x.getAttribute("sql")
					archivo = ""&x.getAttribute("archivo")
					if sql<> "" and archivo<>"" then
						'Response.Write "<br>3) Definida una sentencia SQL y un archivo XML."
						if config_idioma_bd = "" then
							rutaXml = "/" & c_s & "datos/"& session("idioma") &"/"& cualid &"/"&archivo
						else
							rutaXml = "/" & c_s & "datos/"& config_idioma_bd &"/"& cualid &"/"&archivo
						end if
						
						dim re : set re = Server.CreateObject("ADODB.Recordset")
						re.ActiveConnection = conn_
						on error resume next
						
						re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
							
							while not re.eof
							'response.write("-------"&re("R_SECCION"))
'								Response.Write "<br>4) Registro que cumple las condiciones."
								mixml = mixml & "<dato icono="& d & re("R_ICONO") & d &" idseccion="& d & re("R_SECCION") & d &" idsubseccion="& d & re("R_SECCION2") & d &"  id="& d & re("R_ID") & d & " fecha="& d & re("R_FECHA") & d & " hora="& d & re("R_HORA") & d & ">"
								if err=0 then
									for each campo in x.childNodes
										valorCampo = ""&re(campo.getAttribute("campo"))
										if err<>0 then
											unerror = true : msgerror = "No se ha encontrado el campo: >"&campo.getAttribute("campo")&"< en la base de datos. Compruebe que está bien escrito."
											exit for
										end if
										valorCampo = replace(valorCampo,"]]>","]]->")
										mixml = mixml & "<"& campo.getAttribute("att") &"><![CDATA[" & valorCampo & "]]></"& campo.getAttribute("att") &">"
									next
								end if
								mixml = mixml & "</dato>" & vbCrlF
								re.movenext
							wend
							
						on error goto 0
	
						' Muestra todas las secciones, tengan o no registros... las desactivadas no..
						'secci.Source = "select * from secciones where s_registros>0 order by s_orden" : secci.CursorType = 3 : secci.CursorLocation = 2 : secci.LockType = 1 : secci.Open()
						secci.Source = "select * from secciones order by s_orden" : secci.CursorType = 3 : secci.CursorLocation = 2 : secci.LockType = 1 : secci.Open()
						'on error resume next
						if not unerror then
							mixml = mixml & "<secciones>" & vbCrLf
							while not secci.eof
							
								if secci("s_activo")<>false then
								mixml = mixml & "<seccion registros="& d & secci("s_registros") & d & " nombre="& d & secci("s_nombre") & d & " id=" & d & secci("s_id") & d & " foto=" & d & secci("s_foto")  & d & ">"
									
									'Recorro las subsecciones
									subsecci.Source = "select distinct * from secciones2 where S2_ID_S="& secci("s_id")&" order by s2_orden" : subsecci.CursorType = 3 : subsecci.CursorLocation = 2 : subsecci.LockType = 1 : subsecci.Open()
									if not unerror then
											while not subsecci.eof
													if not subsecci.eof then
													mixml = mixml & "<subseccion nombre=" & d & subsecci("S2_NOMBRE") & d &" id=" & d & subsecci("S2_ID") & d &" registros=" & d & subsecci("S2_REGISTROS")  & d & " foto=" & d & subsecci("s2_foto")& d &" >"
													mixml = mixml & "</subseccion>"
													end if
											subsecci.movenext
											wend
									subsecci.close
									end if
								mixml=mixml & "</seccion>" & vbCrLf		
								end if
								secci.movenext
							wend
							mixml = mixml & "</secciones>"  & vbCrLf		
							secci.close
							set secci=nothing		
							mixml = mixml & "</datos>" & vbCrLf

							' Guardar XML
							Set mfso = Server.CreateObject("Scripting.FileSystemObject")
							Set mfsoescribir = mfso.CreateTextFile (Server.MapPath(rutaXml),True)
							mfsoescribir.WriteLine(mixml)
							mfsoescribir.Close
							Set mfsoescribir = Nothing
							Set mfso = Nothing
							Response.Write "<br>XML exportado correctamente ("&cualid &"/"&archivo&")."
						end if
	
						re.close
						set re = nothing
	
					end if ' if sql<> "" then
	
				next
				if unerror then
					Response.Write "Se ha producido errores en la exportación de datos a XML."
				end if
			end if
		end if ' nodoCualid ok
		
	end sub
	
	' Ordena los campos R_orden
	sub reordenaOrdenIdioma()
		' Cadena con los ID Bloqueados
		consultaXOpen "SELECT OI_ID, OI_BLOQUEADA FROM ORDENIDIOMA WHERE OI_BLOQUEADA = 1",1
			if not re.eof then
				n=0
				strBloqueados = "--"
				for n=0 to retotal-1
					strBloqueados = strBloqueados & re("OI_ID") & "-"
					re.movenext
				next
			end if
		consultaXClose()
		
		' ESP
		consultaXOpen "SELECT OI_ORDEN_ESP, OI_BLOQUEADA FROM ORDENIDIOMA ORDER BY OI_ORDEN_ESP ASC",2
		n = 1
		while not re.eof
			if not re("OI_BLOQUEADA") then
				' Salto los Ordenes que estan bloqueados
				while inStr(strBloqueados,"-"&n&"-")>0
					n=n+1
				wend
				re("OI_ORDEN_ESP") = n
			end if
			n = n + 1
			re.MoveNext()
		wend
		consultaXClose()
		
		' ENG
		consultaXOpen "SELECT OI_ORDEN_ENG, OI_BLOQUEADA FROM ORDENIDIOMA ORDER BY OI_ORDEN_ENG ASC",2
		n = 1
		while not re.eof
			if not re("OI_BLOQUEADA") then
				' Salto los Ordenes que estan bloqueados
				while inStr(strBloqueados,"-"&n&"-")>0
					n=n+1
				wend
				re("OI_ORDEN_ENG") = n
			end if
			n = n + 1
			re.MoveNext()
		wend
		consultaXClose()

		' FRA
		consultaXOpen "SELECT OI_ORDEN_FRA, OI_BLOQUEADA FROM ORDENIDIOMA ORDER BY OI_ORDEN_FRA ASC",2
		n = 1
		while not re.eof
			if not re("OI_BLOQUEADA") then
				' Salto los Ordenes que estan bloqueados
				while inStr(strBloqueados,"-"&n&"-")>0
					n=n+1
				wend
				re("OI_ORDEN_FRA") = n
			end if
			n = n + 1
			re.MoveNext()
		wend
		consultaXClose()
		
		' DEU
		consultaXOpen "SELECT OI_ORDEN_DEU, OI_BLOQUEADA FROM ORDENIDIOMA ORDER BY OI_ORDEN_DEU ASC",2
		n = 1
		while not re.eof
			if not re("OI_BLOQUEADA") then
				' Salto los Ordenes que estan bloqueados
				while inStr(strBloqueados,"-"&n&"-")>0
					n=n+1
				wend
				re("OI_ORDEN_DEU") = n
			end if
			n = n + 1
			re.MoveNext()
		wend
		consultaXClose()

		' ITA
		consultaXOpen "SELECT OI_ORDEN_ITA, OI_BLOQUEADA FROM ORDENIDIOMA ORDER BY OI_ORDEN_ITA ASC",2
		n = 1
		while not re.eof
			if not re("OI_BLOQUEADA") then
				' Salto los Ordenes que estan bloqueados
				while inStr(strBloqueados,"-"&n&"-")>0
					n=n+1
				wend
				re("OI_ORDEN_ITA") = n
			end if
			n = n + 1
			re.MoveNext()
		wend
		consultaXClose()
		
	end sub
	
	' Ordena los campos R_orden
	sub reordenaSecciones()
		' Secciones que estan bloqueadas
		consultaXOpen "SELECT S_ORDEN, S_BLOQUEADA FROM SECCIONES WHERE S_BLOQUEADA = 1 ORDER BY S_ORDEN ASC",2
			if not re.eof then
				dim cadena
				cadena = "--"
				while not re.eof
					cadena = cadena & re("S_ORDEN") & "-"
					re.movenext
				wend
			end if
		consultaXClose()

		consultaXOpen "SELECT S_ORDEN, S_BLOQUEADA FROM SECCIONES ORDER BY S_ORDEN ASC",2
		n = 1
		while not re.eof
			if not re("S_BLOQUEADA") then
				' Salto los Ordenes que estan bloqueados
				while inStr(cadena,"-"&n&"-")>0
					n=n+1
				wend
				re("S_ORDEN") = n
			end if
			n = n + 1
			re.MoveNext()
		wend
		consultaXClose()
	end sub

	' Ordena los campos R_orden Secciones 2
	sub reordenaSecciones2(pSeccion)
		' Array con las secciones que estan bloqueadas
		consultaXOpen "SELECT S2_ORDEN, S2_BLOQUEADA FROM SECCIONES2 WHERE S2_ID_S = "& pSeccion &" AND S2_BLOQUEADA = 1 ORDER BY S2_ORDEN ASC",1
			if not re.eof then
				cadena = "--"
				for n=0 to retotal-1
					cadena = cadena & re("S2_ORDEN") & "-"
					re.movenext
				next
			end if
		consultaXClose()
		
		consultaXOpen "SELECT S2_ORDEN, S2_BLOQUEADA FROM SECCIONES2 WHERE S2_ID_S = "& pSeccion &" ORDER BY S2_ORDEN ASC",2
		n = 1
		while not re.eof
			if not re("S2_BLOQUEADA") then
				' Salto los Ordenes que estan bloqueados
				while inStr(cadena,"-"&n&"-")>0
					n=n+1
				wend
				re("S2_ORDEN") = n
			end if
			n = n + 1
			re.MoveNext()
		wend
		consultaXClose()
	end sub
	
	sub upRegSeccion(seccion)
		call exeSql("UPDATE SECCIONES SET S_REGISTROS = S_REGISTROS + 1 WHERE S_ID = "& seccion, conn_)
	end sub
	
	sub downRegSeccion(seccion)
		call exeSql("UPDATE SECCIONES SET S_REGISTROS = S_REGISTROS - 1 WHERE S_ID = "& seccion, conn_)
	end sub

	sub upRegSeccion2(seccion)
		call exeSql("UPDATE SECCIONES2 SET S2_REGISTROS = S2_REGISTROS + 1 WHERE S2_ID = "& seccion, conn_)
	end sub
	
	sub downRegSeccion2(seccion)
		call exeSql("UPDATE SECCIONES2 SET S2_REGISTROS = S2_REGISTROS - 1 WHERE S2_ID = "& seccion, conn_)
	end sub
	
	
	' Aunmentar/disminuir número de subsecciones que contiene un sección
	sub upSeccion(seccion)
		call exeSql("UPDATE SECCIONES SET S_SUBSECCIONES = S_SUBSECCIONES + 1 WHERE S_ID = "& seccion, conn_)
	end sub	
	sub downSeccion(seccion)
		call exeSql("UPDATE SECCIONES SET S_SUBSECCIONES = S_SUBSECCIONES - 1 WHERE S_ID = "& seccion, conn_)
	end sub


	sub traspasoDeSeccion(pAnt,pNue)
		dim ant, nue
		ant = ""&pAnt
		nue = ""&pNue
		if ant <> nue and esNumero(nue) and esNumero(ant) then
			sql = "SELECT S_REGISTROS, S_ID FROM SECCIONES WHERE S_ID = "& ant &" OR S_ID = "& nue
			set re_trS = Server.CreateObject("ADODB.Recordset")
			re_trS.ActiveConnection = conn_
			re_trS.Source = sql : re_trS.CursorType = 3 : re_trS.CursorLocation = 2 : re_trS.LockType = 3 : re_trS.Open()

			if re_trS.recordcount = 2 then
				while not re_trS.eof
					if ""&re_trS("S_ID") = ant then
						re_trS("S_REGISTROS") = re_trS("S_REGISTROS") - 1
					elseif ""&re_trS("S_ID") = nue then
						re_trS("S_REGISTROS") = re_trS("S_REGISTROS") + 1
					end if
					re_trS.MoveNext()
				wend
			end if

			re_trS.Close()
			set re_trS = nothing
		end if
	end sub

	sub traspasoDeSeccion2(pAnt,pNue)
		dim ant, nue
		ant = ""&pAnt
		nue = ""&pNue
		if ant <> nue and esNumero(nue) and esNumero(ant) then
			sql = "SELECT S2_REGISTROS, S2_ID FROM SECCIONES2 WHERE S2_ID = "& ant &" OR S2_ID = "& nue
			set re_trS = Server.CreateObject("ADODB.Recordset")
			re_trS.ActiveConnection = conn_
			re_trS.Source = sql : re_trS.CursorType = 3 : re_trS.CursorLocation = 2 : re_trS.LockType = 3 : re_trS.Open()

			if re_trS.recordcount = 2 then
				while not re_trS.eof
					if ""&re_trS("S2_ID") = ant then
						re_trS("S2_REGISTROS") = re_trS("S2_REGISTROS") - 1
					elseif ""&re_trS("S2_ID") = nue then
						re_trS("S2_REGISTROS") = re_trS("S2_REGISTROS") + 1
					end if
					re_trS.MoveNext()
				wend
			end if

			re_trS.Close()
			set re_trS = nothing
		end if
	end sub

%>