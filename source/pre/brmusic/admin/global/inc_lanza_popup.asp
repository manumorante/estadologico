<%

dim conn_
conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("/" & c_s & "datos/"& idioma &"/popup/popup.mdb")

sub lanzaPopUp

	if ""&idioma <> "" then
		' *** CARGA DEL XML DE CONFIGURACIÓN ***
		xmlAdmin = "/" & c_s & "datos/xml_admin_config.xml"
		set xml_config = CreateObject("MSXML.DOMDocument")
		if not xml_config.Load(server.MapPath(xmlAdmin)) then
			popupError = true : popupMsgError = "'XML config' Error de carga."
		else
			set nodoConfig = xml_config.selectSingleNode("configuracion")
			if typeName(nodoConfig) = "Nothing" or typeName(nodoConfig) = "Empty" then
				popupError = true : popupMsgError = "'XML config' Error de estructura."
			end if
		end if
		
		if not popupError then
			set nodoCualid = xml_config.selectSingleNode("configuracion/popup")
			if typeName(nodoCualid) = "Nothing" or typeName(nodoCualid) = "Empty" then
				popupError = true : popupMsgError = "No hay popUp definidos en el XML."
			end if
		end if	
	
		if not popupError then
			set nodoVisor = xml_config.selectSingleNode("configuracion/popup/visor")
			if typeName(nodoVisor) = "Nothing" or typeName(nodoVisor) = "Empty" then
				popupError = true : popupMsgError = "No se ha encontrado en nodo de configuración para el visor."
			end if
		end if
	
		if not popupError then
			' Abro la conexion a la base de datos
			sql = "SELECT * FROM REGISTROS ORDER BY R_ORDEN DESC"
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conn_ : re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 2 : re.Open()
			if not re.eof then
				while not re.eof
				
					nav_titulo = re("R_TITULO")
					nav_activo = cbool(re("R_ACTIVO"))
					nav_fechaini = CDate(re("R_FECHAINI"))
					nav_fechafin = CDate(re("R_FECHAFIN"))
					nav_foto = ""&re("R_FOTO")
					nav_ancho = re("R_TEXT1")
					nav_alto = re("R_TEXT2")
					id = re("R_ID")
					
					if ""&nav_ancho="" then
						nav_ancho = 200
					end if
		
					if ""&nav_alto = "" then
						nav_alto = 200
					end if
					
					if ""&nav_fechaini <> "" and nav_fechaini <> "0:00:00" then
						nav_fechaini = CDate(nav_fechaini)
					else
						nav_fechaini = ""
					end if
		
					if ""&nav_fechafin <> "" and nav_fechafin <> "0:00:00" then
						nav_fichafin = CDate(nav_fechafin)
					else
						nav_fechafin = ""
					end if
					
					pop = false
					hoy = Date()
					
					if nav_fechaini <> "" and nav_fechafin <> "" then
						if hoy >= nav_fechaini and hoy < nav_fechafin then
							pop = true
							if not nav_activo then
								call setActivo(id,1)
							end if
						end if
					elseif nav_fechaini <> "" and nav_fechafin = "" then
						if hoy >= nav_fechaini then
							pop = true
							if not nav_activo then
								call setActivo(id,1)
							end if
						end if
					elseif nav_fechafin <> "" and nav_fechaini = "" then
						if hoy < nav_fechafin and nav_activo then
							pop = true
						elseif hoy >= nav_fechafin and nav_activo then
							' Si se ha pasado la fecha (o es hoy el dia) y estÁ acitvo: lo desactivamos (y no sale)
							call setActivo(id,0)
						end if
					else
						' No hay restricciones de fechas
						if nav_activo then
							pop = true
						end if
					end if
					
					' Si se ha elegido "tamaño igual a foto" iniciamos la ventana en tamaño estandar 100x100
					if ""&re("R_OPCION1") = "si" then
						nav_ancho = 100
						nav_alto = 100
					end if
					
					' Almaceno los datos del registro para popUp con ellos en caso de ser valido (por que el proximo registro puede no ser POP y no quiero que me sobreescriba los datos)
					if pop then
						pop_id = id
						pop_titulo = nav_titulo
						pop_ancho = nav_ancho
						pop_alto = nav_alto
						pop_pop = true
					end if

					re.movenext
				wend
			else
				popupError = true : popupMsgError = "No se ha encontrado ningún registro."
			end if
			re.close
			set re = nothing
		end if
		if not popupError then
			if pop_pop then%>
			<script language="javascript" type="text/javascript">
				function lanzapo_win(theURL,winName,ancho,alto,barras) { 
					try{
						var winl = (screen.width - ancho) / 2;
						var wint = (screen.height - alto) / 2;
						var paramet='top='+wint+',left='+winl+',width='+ancho+',height='+alto+',scrollbars='+barras+'';
						var splashWin=window.open(theURL,winName,paramet);
						splashWin.focus();
					}catch(unerror){}
				}
				lanzapo_win("../admin/global/popup.asp?id=<%=pop_id%>&titulo=<%=pop_titulo%>&idioma=<%=idioma%>","PopUp",<%=pop_ancho%>,<%=pop_alto%>,0)
			</script>
			<%else
				' Las fechas dicen que no le toca salir
			end if
		else%>
			<script>//alert("<%=popupMsgError%>")</script>
			<!-- <%=popupMsgError%> -->
		<%end if
	end if
end sub


function setActivo (id,activo)
	if ""&id<>"" and ""&activo <> "" then
		call exeSQL("UPDATE REGISTROS SET R_ACTIVO = "& activo &" WHERE R_ID = " & id,conn_)
	end if
end function


' Reseteo los mensajes de errror
popupError = false
popupMsgError = ""

%>