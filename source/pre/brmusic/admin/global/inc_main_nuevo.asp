<%

		seccion = numero(request("seccion"))
		seccion2 = numero(request("seccion2"))
		seccion3 = numero(request("seccion3"))
		if seccion2 <=0 then
			seccion2 = 1
		end if
		if seccion3 <=0 then
			seccion3 = 1
		end if
		
		id = ""&request.Form("id")
		fecha = date()
		fechaini = date()
		fechafin = date()
		
		if cualid = "usuarios" and seccion >0 then
			set miGrupo = setGrupo(seccion)
			set miGrupoConfig = miGrupo.selectSingleNode("config")
		end if
	
		' Duplicación de registros
		' ----------------------------------------------------------------------------
		duplicar = false
		if request.Form("comun") <> "" and isNumeric(id) and id <> "" then
			duplicar = true
			sql = "SELECT * FROM REGISTROS WHERE R_ID = "&id
			dim dupli
			set dupli = Server.CreateObject("ADODB.Recordset")
			dupli.ActiveConnection = conn_
			dupli.Source = sql : dupli.CursorType = 3 : dupli.CursorLocation = 2 : dupli.LockType = 3 : dupli.Open()
			retotal = dupli.recordcount
			if not dupli.eof then
				duplicar = true
				titulo = ""&dupli("R_TITULO")
				fuente = ""&dupli("R_FUENTE")
				enlace = ""&dupli("R_ENLACE")
				seccion = ""&dupli("R_SECCION")
				fecha = dupli("R_FECHA")
				hora = dupli("R_HORA")
				fechaini = dupli("R_FECHAINI")
				fechafin = dupli("R_FECHAFIN")
			end if
		end if
		

	' Inserción
	' ------------------------------------

	if request.Form("insertar") <> "" then
		enportada = 0 + ("0"&request.Form("enportada"))
		activo = 0 + ("0"&request.Form("activo"))
		titulo = replace(""&request.Form("titulo"),"'","''")
		email = replace(""&request.Form("email"),"'","''")
		fuente = replace(""&request.Form("fuente"),"'","''")
		enlace = replace(lcase(""&trim(request.Form("enlace"))),"'","''")
		ordenidioma = replace(""&request.Form("orden_idioma"),"'","''")
		piefoto = left(replace(""&request.Form("piefoto"),"'","''"),254)
		pos_foto = request.Form("pos_foto")
		pos_icono = request.Form("pos_icono")
		
		if cualid = "usuarios" then
			clave = replace(""&request.Form("clave"),"'","''")
			clave_r = replace(""&request.Form("clave_r"),"'","''")
			claveCodi = sha256(clave)

			if clave <> clave_r then
				errorform = true : msgerrorform = msgerrorform & "<br>Las claves no coinciden."
			end if
	
			if clave = "" then
				errorform = true : msgerrorform = msgerrorform & "<br>Debe escribir una clave."
			end if
		end if

		if inStr(enlace,"http://") <= 0 and enlace <> "" then
			enlace = "http://"&enlace
		end if

		if request.Form("fecha") = 1 and  ""&request.Form("fecha_dia") <> "" and ""&request.Form("fecha_mes") <> "" and ""&request.Form("fecha_ano") <> "" then
			fecha = request.Form("fecha_dia")& "/" &request.Form("fecha_mes")& "/" &request.Form("fecha_ano")
		else
			fecha = "0:00:00"
		end if
		
		if ""&request.Form("hora_hora") <> "" and ""&request.Form("hora_minutos")<> "" then
			hora = Right("00"&request.Form("hora_hora"),2)& ":" &Right("00"&request.Form("hora_minutos"),2)
		else
			hora = "0:00"
		end if
		
		if request.Form("fechaini") = 1 and  ""&request.Form("fechaini_dia") <> "" and ""&request.Form("fechaini_mes") <> "" and ""&request.Form("fechaini_ano") <> "" then
			fechaini = request.Form("fechaini_dia")& "/" &request.Form("fechaini_mes")& "/" &request.Form("fechaini_ano")
		else
			fechaini = "0:00:00"
		end if
		
		if request.Form("fechafin") = 1 and  ""&request.Form("fechafin_dia") <> "" and ""&request.Form("fechafin_mes") <> "" and ""&request.Form("fechafin_ano") <> "" then
			fechafin = request.Form("fechafin_dia")& "/" &request.Form("fechafin_mes")& "/" &request.Form("fechafin_ano")
		else
			fechafin = "0:00:00"
		end if
		
		if not esNumero(seccion) then
			errorform = true : msgerrorform = msgerrorform & "<br>Escoja una sección."
		end if

		if titulo = "" then
			errorform = true : msgerrorform = msgerrorform & "<br>Escriba un titulo."
		end if
		
		if ""&request.Form("nuevosalfinal") = "1" then
			orden = "9999999"
		else
			orden = "0.9"
		end if
		
		if not errorform then
			sql = "INSERT INTO"

			Dim sql_nombres, sql_valores

			' Campos fijos
			' ----------------------------------

			sql_nombres = sql_nombres & "R_TITULO"
			sql_valores = sql_valores & "'"& titulo &"'"

			' Campos para idiomas
			if inStr(config_str_idiomas,"|eng|") then
				sql_nombres = sql_nombres & ", R_TITULO_ENG"
				sql_valores = sql_valores & ",'"& titulo &"'"
			end if

			sql_nombres = sql_nombres & ", R_CLAVE, R_EMAIL, R_SECCION, R_USUARIO, R_FUENTE, R_ORDEN, R_ORDEN_SECCION, R_ORDEN_SECCION2, R_PORTADA, R_ACTIVO, R_ENLACE, R_FECHA, R_HORA, R_FECHAINI, R_FECHAFIN, R_PIE_FOTO, R_POS_FOTO, R_POS_ICONO"
			sql_valores = sql_valores & ", '"& claveCodi &"', '"& email &"', "& seccion &", "& session("usuario") &",'"& fuente &"', "& orden &", "& orden &", "& orden &", '"& enportada &"', '"& activo &"', '"& enlace &"', '"& fecha&"', '"& hora &"', '"& fechaini &"', '"& fechafin &"', '"& piefoto &"', '"& pos_foto &"', '"& pos_icono &"'"

			if config_activo_seccion2 then
				sql_nombres = sql_nombres & ", R_SECCION2"
				sql_valores = sql_valores & ", "& seccion2
			end if
			
			' Utilizo el nodo miGrupo que contiene seteado el grupo de la seccion escogida
			if cualid = "usuarios" and typeOK(miGrupo) then
				set nodosCampo = miGrupo.selectNodes("//grupos/grupo[@id="& seccion &"]//dato")
			else
				set nodosCampo = nodoCualid.childNodes
			end if

			' Campos configurables
			for each a in nodosCampo
				if a.nodeName = "dato" then
					valor = replace(""&request.Form(""&a.getAttribute("nombrecorto")),"'","''")
					campo = Ucase(""&a.getAttribute("campo"))
					if campo <> "" then
						if ""&a.getAttribute("requerido") = "1" and ""&valor="" then
							errorform = true : msgerrorform = msgerrorform & "<br>El campo """& a.getAttribute("titulo") &""" es obligatorio."
						end if

						sql_nombres = sql_nombres & ", R_"& campo &""
						sql_valores = sql_valores & ", '"& valor &"'"

						if inStr(config_str_idiomas,"|eng|") then
							sql_nombres = sql_nombres & ", R_"& campo &"_ENG" 
							sql_valores = sql_valores & ", '"& valor &"'"
						end if

					end if
				end if
			next

			sql = sql & " REGISTROS ("& sql_nombres &")"
			sql = sql & " VALUES ("& sql_valores &")"

			if not errorform  then
				set conn_activa = server.CreateObject("ADODB.Connection")
				conn_activa.Open conn_
				conn_activa.execute sql
				
				if err=0 then


					' Incrementar contador de la sección y subsección
					upRegSeccion(seccion)
					if config_activo_seccion2 then
						upRegSeccion2(seccion2)
					end if

					reOrdena()

					mi_seccion = seccion
					reOrdena()
			
					mi_seccion2 = seccion2
					reOrdena()
					
					' Libero conexión activa.
					conn_activa.Close
					set conn_activa = nothing

					' Insertar secciones en otras cualidades.
					'-------------------------------------------------
					relacion = ""& request.Form("relacion")
					if relacion <> "" or config_infoemail then

						' Localizar la última ID (el registro recien insertado).
						id = 0
						titulo = ""
						sql = "SELECT TOP 1 R_ID, R_TITULO FROM REGISTROS ORDER BY R_ID DESC"
						set re = Server.CreateObject("ADODB.Recordset")
						re.ActiveConnection = conn_
						re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 1
						re.Open()
						id = re("R_ID")
						titulo = re("R_TITULO")
						re.Close()
						set re = nothing
					end if
						
					if relacion <> "" then
						' Insertamos una nueva seccion en la cualidad establecida
						if id > 0 then
							conn_temp = conn_
							conn_ = replace(conn_,"\"& cualid &"\","\"& relacion &"\")
							conn_ = replace(conn_,cualid &".mdb",relacion &".mdb")
							insertarSeccion titulo, 1, id, 0, 0, 1
							conn_ = conn_temp
						end if

					end if


					%>

					Un momento ...
					<script language="javascript" type="text/javascript">
					
					try{
						// Refrescar frame izquierda
						var f = top.frames[1].frames[0].f // Frame de la izquierda
						f.ac.value = ""
						f.target = ""
						f.submit()
					}catch(unerror){}
					<%
					' popMail '
					if config_infoemail then
						infoemail = "&infoemail=1"
					end if

					if Request.Form("foto") = 1 then
						consultaXOpen "SELECT R_ID FROM REGISTROS ORDER BY R_ID DESC",1%>
						location.href='archivos_frames.asp?ac=formguardarfoto&id=<%=re("R_ID")%>&icono=<%=Request.Form("icono")%>&archivo=<%=Request.Form("archivo")%><%=infoemail%>'
						<%consultaXClose()
					elseif ""&Request.Form("icono") = "1" then
						consultaXOpen "SELECT R_ID FROM REGISTROS ORDER BY R_ID DESC",1%>
						location.href='archivos_frames.asp?ac=formguardaricono&id=<%=re("R_ID")%>&archivo=<%=Request.Form("archivo")%><%=infoemail%>'
						<%consultaXClose()
					elseif Request.Form("archivo") = 1 then
						consultaXOpen "SELECT R_ID FROM REGISTROS ORDER BY R_ID DESC",1%>
						location.href='archivos_frames.asp?ac=formguardararchivo&id=<%=re("R_ID")%><%=infoemail%>'
						<%consultaXClose()
					else

						' Informar por email.'
						' ------------------------------------------------------------------------ '
						if config_infoemail then
							' Nombre del usuario '
							consultaXOpen "SELECT S_NOMBRE FROM SECCIONES WHERE S_ID = "& seccion,1
							if not re.eof then
								nombre_seccion = re("S_NOMBRE")
							end if
							consultaXClose()
	
							' Leemos el usuario a partir del nombre de usuario'
							dim ru, ruTotal
							consultaUsuarios "SELECT R_TITULO, R_EMAIL FROM REGISTROS WHERE R_TITULO = '"& nombre_seccion &"'",1
							if ruTotal >0 then
								if not ru.eof then%>
									// nombre/sección, email, asunto, cuerpo
									popMail('<%=ru("R_TITULO")%>','<%=ru("R_EMAIL")%>','Nueva tarea asignada','Se le ha asignado una nueva tarea.',<%=id%>)
								<%end if
							end if
							consultaUsuariosClose()
						end if
						' ------------------------------------------------------------------------ '

						consultaXOpen "SELECT R_ID, R_SECCION FROM REGISTROS ORDER BY R_ID DESC",1
						if re("R_SECCION") = 2 then%>
							location.href='inicio.asp?ac=editarpermisos&id=<%=re("R_ID")%>';
						<%else%>
							location.href='inicio.asp';
						<%end if%>
					<%end if%>
					</script>
				<%
				else
					if inStr(err.description,"R_") then
						Response.Redirect("inicio.asp?msg=<b>Configuración errónea</b><br>Parece que el XML de configuración no esta correcto.<br>Por favor, consúltelo con su administrador.<!-- "&err.description&" -->")
					else
						Response.Redirect("inicio.asp?msg=<b>Error</b><br>"& err.description)
					end if
				end if
				on error goto 0
			end if
		end if
	end if  

	if config_activo_seccion2 then
		Response.Write "<script language='javascript' type='text/javascript'>" & vbCrlf
		' Leo las subsecciones y las meto en un array
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = "SELECT * FROM SECCIONES2 ORDER BY S2_ID_S, S2_ORDEN"
		re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
	
			a=0
			b=0
			while not re.eof
				if ""&estaSecc <> ""&re("S2_ID_S") then
					b = b + 1
					arrTit = "arrTituloSeccion"& re("S2_ID_S")
					arrVal = "arrValorSeccion"& re("S2_ID_S")
					Response.Write "var "& arrTit &" = new Array()" & vbCrlf
					Response.Write "var "& arrVal &" = new Array()" & vbCrlf
				end if
				Response.Write arrTit&"["&a&"] = '" & re("S2_NOMBRE") &"'"& vbCrlf
				Response.Write arrVal&"["&a&"] = '" & re("S2_ID") &"'"& vbCrlf
				a=a+1
				estaSecc = re("S2_ID_S")
				re.movenext
			wend
		
		re.Close()
		set re = nothing  
		Response.Write "</script>"
	end if
  %>

<script language="javascript" type="text/javascript">
		try{
			f.ac.value = "nuevo"
		}catch(unerror){}
	
	function cargaSubSecciones(c){
		try{
			borrarCombo(f.seccion2)
			var arrT = eval("arrTituloSeccion"+c.value)
			var arrV = eval("arrValorSeccion"+c.value)
			for(var n=0; n<arrT.length; n++){
				if (""+arrT[n] != "" && ""+arrV[n] != "" && ""+arrT[n] != "undefined" && ""+arrV[n] != "undefined"){
					opt = new Option(arrT[n], arrV[n]);
					f.seccion2.options[f.seccion2.length] = opt;
				}
			}
		}catch(unerror){}
	}
    function borrarCombo(combo){
	    for(var c=combo.length;c>0;c--){
            combo.options[c] = null;
        }
        combo.options[0].selected = true;
    }
</script>

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
		
      <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
		
      <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Nuevo <%if duplicar then%> a partir de existente<%end if%></font></b></td>
      <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
    </tr>
	</table>
	<input type="hidden" name="nuevosalfinal" value="<%=request.Form("nuevosalfinal")%>">
	<%if errorform then%>
	<br>
		<div align="center" class="Estilo4"><%=msgerrorform%> </div>
	<%end if%>
  <br>
	<%
	' Posible relación del grupo seleccionado con otra cualidad.
	if typeOK(miGrupoConfig) then%>
		<input type="hidden" name="relacion" value="<%=""& miGrupoConfig.getAttribute("relacion")%>">
	<%end if%>
  <table width="100%"  border="0" cellspacing="0" cellpadding="2">
    <tr>
      <td><b>Secci&oacute;n:</b>
	  <%consultaXOpen "SELECT * FROM SECCIONES ORDER BY S_ORDEN",1%>
	  <select name="seccion" class="campoAdmin" onChange="cargaSubSecciones(this)">
	  <%if reTotal > 1 then%>
        <option>Secci&oacute;n ...</option>
		<%end if%>
	  <%while not re.eof%>
        <option value="<%=re("S_ID")%>" <%if strComp(seccion,re("S_ID"))=0 then Response.Write "selected" end if%>><%=re("S_NOMBRE")%></option>
		<%re.movenext : wend%>
      </select>
	  <%consultaXClose()
	  if config_activo_seccion2 then
		  consultaXOpen "SELECT * FROM SECCIONES2 WHERE S2_ID_S = "& seccion &" ORDER BY S2_ORDEN",1
		  if not re.eof then%>
		   /
		  <select name="seccion2" class="campoAdmin">
			<option value="1">Sub secci&oacute;n ...</option>
		  <%while not re.eof%>
			<option value="<%=re("S2_ID")%>" <%if strComp(seccion2,re("S2_ID"))=0 then Response.Write "selected" end if%>><%=re("S2_NOMBRE")%></option>
			<%re.movenext : wend%>
		  </select>
		  <%end if
		  consultaXClose()
	 end if%></td>
	  
      <td align="right">

	  <%if config_portada then%>
	  <table  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><input name="enportada" type="checkbox" id="enportada" value="1" <%if ""&request.Form("portada") = "1" then Response.Write "checked" end if%>></td>
          <td><img src="img/bandera.gif" width="18" height="18"></td>
          <td><label for="enportada"><b>En portada </b></label></td>
        </tr>
      </table>
	  <%end if%>

	  <%if config_activo then%>
	  <table  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><input name="activo" type="checkbox" id="activo" value="1" <%if ""&request.Form("activo") = "1" then Response.Write "checked" end if%>></td>
          <td><label for="activo"><b>Activo</b></label></td>
        </tr>
      </table>
	  <%end if%>	  

	  </td>
    </tr>
	
	<%if config_fechainifin then%>
<tr bgcolor="#FFFFFF">
      <td colspan="2">
	    <b>&nbsp;<%=config_nom_fechaini%></b>
	    <input name="fechaini" id="fechaini_si" type="radio" onClick="siAnadirFechaini()" value="1" <%if duplicar and ""&fechaini <> "0:00:00" then Response.Write "checked" end if%>><label for="fechaini_si">S&iacute;&nbsp;</label>
		<select name="fechaini_dia" <%if not duplicar or not ""&fechaini <> "0:00:00" then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechaini_dia">
          <% for n=1 to 31 %>
          <option value="<%=n%>" <%if n = day(fechaini) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
        <select name="fechaini_mes" <%if not duplicar or not ""&fechaini <> "0:00:00"then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechaini_mes">
          <% for n=1 to 12 %>
          <option value="<%=n%>"<%if n = month(fechaini) then Response.Write "selected" end if%>><%=arrMes(n-1)%></option>
          <% next %>
        </select>
        <select name="fechaini_ano" <%if not duplicar or not ""&fechaini <> "0:00:00"then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechaini_ano">
          <% for n=Year(date())-12 to Year(date())+1%>
          <option value="<%=n%>"<%if n = year(fechaini) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
	  <input name="fechaini" id="fechaini_no" type="radio" onClick="noAnadirFechaini()" value="0" <%if not duplicar or ""&fechaini = "0:00:00" then Response.Write "checked" end if%>><label for="fechaini_no">No</label></td>
    </tr>
	
<tr bgcolor="#FFFFFF">
      <td colspan="2">
	    <b>&nbsp;<%=config_nom_fechafin%></b>
	    <input name="fechafin" id="fechafin_si" type="radio" onClick="siAnadirFechafin()" value="1" <%if duplicar and ""&fechaini <> "0:00:00" then Response.Write "checked" end if%>><label for="fechafin_si">S&iacute;&nbsp;</label>
		<select name="fechafin_dia" <%if not duplicar or not ""&fechafin <> "0:00:00" then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechafin_dia">
          <% for n=1 to 31 %>
          <option value="<%=n%>" <%if n = day(fechafin) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
        <select name="fechafin_mes" <%if not duplicar or not ""&fechafin <> "0:00:00"then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechafin_mes">
          <% for n=1 to 12 %>
          <option value="<%=n%>"<%if n = month(fechafin) then Response.Write "selected" end if%>><%=arrMes(n-1)%></option>
          <% next %>
        </select>
        <select name="fechafin_ano" <%if not duplicar or not ""&fechafin <> "0:00:00"then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechafin_ano">
          <% for n=Year(date())-12 to Year(date())+1%>
          <option value="<%=n%>"<%if n = year(fechafin) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
	  <input name="fechafin" id="fechafin_no" type="radio" onClick="noAnadirFechafin()" value="0" <%if not duplicar or ""&fechafin = "0:00:00" then Response.Write "checked" end if%>><label for="fechafin_no">No</label></td>
    </tr>
	<%end if
	
	if config_fecha >0 then%>
    <tr>
      <td colspan="2">
	    <b><%=config_nom_fecha%></b>
		<%if config_fecha =1 then%>
	    <input name="fecha" type="radio" onClick="siAnadirFecha()" value="1" <%if duplicar and ""&fecha <> "0:00:00" then Response.Write "checked" end if%>>
	    S&iacute;
		<%else%>
		<input type="hidden" name="fecha" value="1">
		<%end if%>
	    <select name="fecha_dia" <%if (not duplicar or not ""&fecha <> "0:00:00") and config_fecha<>2 then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fecha_dia">
          <% for n=1 to 31 %>
          <option value="<%=n%>" <%if n = day(fecha) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
        <select name="fecha_mes" <%if (not duplicar or not ""&fecha <> "0:00:00") and config_fecha<>2then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fecha_mes">
          <% for n=1 to 12 %>
          <option value="<%=n%>"<%if n = month(fecha) then Response.Write "selected" end if%>><%=arrMes(n-1)%></option>
          <% next %>
        </select>
        <select name="fecha_ano" <%if (not duplicar or not ""&fecha <> "0:00:00") and config_fecha<>2then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fecha_ano">
          <% for n=Year(date())-12 to Year(date())+1%>
          <option value="<%=n%>"<%if n = year(fecha) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
		<%if config_fecha =1 then%>
	  <input name="fecha" type="radio" onClick="noAnadirFecha()" value="0" <%if not duplicar or ""&fecha = "0:00:00" then Response.Write "checked" end if%>>      
	  No
	  <%end if%></td>
    </tr>
	<%end if
if config_hora then%>
    <tr>
      <td colspan="2">
	    <b>Hora</b>
	    <select name="hora_hora" class="campoAdmin" id="hora_hora">
          <% for n=0 to 23 %>
          <option value="<%=n%>" ><%=Right("00"&n,2)%></option>
          <% next %>
        </select>
        <strong>: </strong>
        <select name="hora_minutos" class="campoAdmin" id="hora_minutos">
          <% for n=0 to 11 %>
          <option value="<%=n*5%>"><%=Right("00"&n*5,2)%></option>
          <% next %>
        </select>
	</td>
    </tr>
	<%end if%>	
	
	
    <tr>
      <td colspan="2"><b><%=config_nom_titulo%></b>*</td>
    </tr>
    <tr>
      <td colspan="2"><input name="titulo" type="text" class="campoAdmin" id="titulo" style="width:100%" value="<%=titulo%>" maxlength="255"></td>
    </tr>
	<%if ""&cualid = "usuarios" then%>
    <tr>
      <td colspan="2"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
        <tr>
            <td width="50%"><b>Clave</b></td>
            <td width="50%"><b>Repetir clave*</b></td>
        </tr>
          <tr>
            <td width="50%"><input name="clave" type="password" class="campoAdmin" id="clave" maxlength="255"></td>
            <td width="50%"><input name="clave_r" type="password" class="campoAdmin" id="clave_r" maxlength="255"></td>
        </tr>
        </table></td>
    </tr>
		<tr>
		  <td colspan="2"><b>E-mail</b></td>
    </tr>
		<tr>
		  <td colspan="2"><input name="email" type="text" class="campoAdmin" id="email" style="width:100%" maxlength="255"></td>
    </tr>
	<%end if%>

<!-- Campo configurables -->
<%

	if config_ordenidioma then
		sql = "SELECT * FROM ORDENIDIOMA"
		consultaXOpen sql,1
	end if
	


' Utilizo el nodo miGrupo que contiene seteado el grupo de la seccion escogida
if cualid = "usuarios" then
	if typeOK(miGrupo) then
		set nodosCampo = miGrupo.selectNodes("//grupos/grupo[@id="& seccion &"]//dato")
	else
		set nodosCampo = nodoCualid.childNodes
		%>
		<tr><td colspan="2"><font color="#FF0000"><b>No se ha definido el grupo '<%=seccion%>' en el XML de grupos.</b></font></td></tr>
		<tr><td colspan="2"><font color="#FF0000">El mapeo de datos se está realizando desde la cualidad.</font></td>
		</tr>
		<%
	end if
else
	set nodosCampo = nodoCualid.childNodes
end if

numCampo = 0
for each a in nodosCampo
	if a.nodeName = "dato" then
		c_titulo = ""&a.getAttribute("titulo")
		c_nombre = ""&a.getAttribute("nombrecorto")
		c_tipo = ""&a.getAttribute("tipo")
		campo = ""&a.getAttribute("campo")
		if campo = "" then
			if c_nombre <> "usuario" and c_nombre <> "email" and c_nombre <> "clave" and c_nombre <> "clave_r" then
				%>
				<tr><td colspan="2"><font color="#FF0000"><%
				if c_titulo <> "" then
					Response.Write "<b>"& c_titulo &"</b>:<br>"
				else
					if ""&a.getAttribute("nombre") <> "" then%>
						<b>Configuración antigua:</b> (<%=a.getAttribute("nombre")%>)<br>
						&middot; El atributo <i>nombre</i> debe llamarse <i>título</i>.
						<br>
					<%end if
				end if
				%> 
					&middot; No se ha indicado un campo para la base de datos.</font></td>
				</tr>
				<%
			end if
		else
			numCampo = numCampo + 1
			if duplicar then
				valor = ""&dupli("R_"&c_nombre)
			else
				valor = request.Form(c_nombre)
			end if
			%>

    <tr>
      <td colspan="2"><b><%=c_titulo%></b><%if ""&a.getAttribute("requerido")="1" then%>*<%end if%></td>
    </tr>

    <tr>
      <td colspan="2" valign="top">
	  <%

	  select case c_tipo
	  
case "texto"%>
	  <input name="<%=c_nombre%>" type="text" class="campoAdmin" id="subtitulo" style="width:100%" value="<%=valor%>" maxlength="255">

	  <%case "memo"%>
	  <textarea name="<%=c_nombre%>" rows="<%=a.getAttribute("filas")%>" wrap="virtual" class="areaAdmin" style="width:100%"><%=valor%></textarea>
	  <%if a.getAttribute("editorhtml") = 1 then%>
	  <script language="javascript1.2">
	  editor_generate("<%=c_nombre%>");
	  </script>
	  <%end if%>

	  <%case "combo"%>
	  <select name="<%=c_nombre%>" class="campoAdmin">
		<option value="">Seleccione una ...</option>
		<%for each opcion in a.childNodes%>
		<option value="<%=opcion.getAttribute("valor")%>" <%
		if valor = opcion.getAttribute("valor") then
			Response.Write "selected"
		elseif opcion.getAttribute("default")=1 then
			Response.Write "selected"
		end if%>><%=opcion.getAttribute("titulo")%></option>
		<%next%>
	</select>

	  <%case "opcion"
	  
	  	inc = 0
		for each opcion in a.childNodes
		inc = inc + 1%>
		<input type="radio" name="<%=c_nombre%>" id="<%=c_nombre&inc%>" value="<%=opcion.getAttribute("valor")%>" <%if valor = "" and ""&opcion.getAttribute("default") = "1" then Response.Write "checked" else if valor = opcion.getAttribute("valor") then Response.Write "checked" end if end if%>><label for="<%=c_nombre&inc%>"><%=opcion.getAttribute("titulo")%></label>
		<%next%>
		
	  <%case "check"
	  
	  	valores = split(""&valor,",")
		for each valor in valores
			cadena = cadena &"|"& trim(valor)
		next
		cadena = cadena &"|"

	  	inc = 0
		for each opcion in a.childNodes
		inc = inc + 1
		%>
		<input  type="checkbox" name="<%=c_nombre%>" id="<%=c_nombre&inc%>" value="<%=opcion.getAttribute("valor")%>" <%if inStr(cadena,opcion.getAttribute("valor"))>0 then Response.Write "checked" end if%>><label for="<%=c_nombre&inc%>"><%=opcion.getAttribute("titulo")%></label>
		<%next%>
		
	<%case "color"%>
		<div align="center">
		<input name="<%=c_nombre%>" type="text" class="campoAdmin" value="<%=valor%>" size="8" maxlength="7" readonly="true">
		<div id="capaColorFlash" style="position:absolute; z-index:<%=10-numCampo%>;width: 172; height: 155; visibility: visible;">
		<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="172" height="155">
		<param name="movie" value="color.swf?nombre=<%=c_nombre%>">
		<param name="quality" value="high">
		<param name="WMODE" value="transparent">
		<embed src="color.swf?nombre=<%=c_nombre%>" width="172" height="155" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent"></embed>
		</object>
		</div>
		<br>
		</div>
	<%case "orden_idioma"
	
		if config_ordenidioma then%>

			<select name="orden_idioma" class="campoAdmin">
				<option value="">Seleccione una ...</option>
				<% while not re.eof%>
				<option value="<%=re("OI_TITULO")%>"><%=re("OI_TITULO")%></option>
				<%re.movenext : wend%>
			</select>
			
		<%end if
	
	case else

	if c_tipo = "" then%>
		<font color="#FF0000"><b>No se ha difinido un tipo para este campo.</b></font>
	<%else%>
		<font color="#FF0000"><b>El tipo de campo "<%=c_tipo%>" no es v&aacute;lido.</b></font>
	<%end if
end select

%></td>
    </tr>
<%
		end if ' campo <> ""
	end if ' nodeName = dato
	
next

if config_ordenidioma then
	consultaXClose()
end if

%>

<!-- FIN: Campo configurables -->



	
<%if config_fuente then%>
    <tr>
      <td colspan="2" valign="top"><b><%=config_nom_fuente%></b></td>
    </tr>
    <tr>
      <td colspan="2" valign="top"><input name="fuente" type="text" class="campoAdmin" id="fuente" style="width:100%" value="<%=fuente%>" maxlength="255"></td>
    </tr>
    <tr>
      <td colspan="2" valign="top"><b><%=config_nom_enlace%></b></td>
    </tr>
    <tr align="left">
      <td colspan="2"><input name="enlace" type="text" class="campoAdmin" id="enlace" style="width:100%" value="<%=request.Form("enlace")%>" maxlength="255"></td>
    </tr>
<%end if%>
<%if config_foto then%>
    <tr>
      <td><b><%=config_nom_foto%></b></td>
      <td>
	  <%if config_posicion_foto then%>
	  <b>Posici&oacute;n</b>
	  <%end if%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>
	  <input name="foto" id="foto_si" type="radio" value="1"><label for="foto_si">A&ntilde;adir</label>
	  <input name="foto" id="foto_no" type="radio" value="0" checked><label for="foto_no">No a&ntilde;adir</label></td>
	  <td>
	  <%if config_posicion_foto then%>
	  <input name="pos_foto" id="pos_foto_izq" type="radio" value="izq">
        <label for="pos_foto_izq">Izquierda</label>
        <input name="pos_foto" type="radio" id="pos_foto_der" value="der" checked>
        <label for="pos_foto_der">Derecha</label>
		<%end if%></td>
    </tr>

	<%if config_pie_foto then%>
    <tr>
      <td colspan="2"><b>Pie de foto:</b></td>
    </tr>
    <tr>
      <td colspan="2"><textarea name="piefoto" cols="" rows="3" wrap="virtual" class="areaAdmin" id="piefoto" style="width:100%"></textarea></td>
    </tr>
	<%end if%>

<%end if%>

<%if config_icono then%>
    <tr>
      <td><b><%=config_nom_icono%></b></td>
      <td>
	  <%if config_posicion_icono then%>
	  <b>Posici&oacute;n</b>
	  <%end if%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>
	  <%if config_iexplorer then%>
	  <input name="icono" type="radio" id="icono_auto" value="auto">
	  <label for="icono_auto">Auto</label>
	  <%end if%>
	  <input name="icono" id="icono_si" type="radio" value="1"><label for="icono_si">A&ntilde;adir</label>
	  <input name="icono" type="radio" id="icono_no" value="0" checked>
	  <label for="icono_no">No a&ntilde;adir</label></td>
      <td>
	  <%if config_posicion_icono then%>
	  <input name="pos_icono" id="pos_icono_izq" type="radio" value="izq">
	  <label for="pos_icono_izq">Izquierda</label>
      <input name="pos_icono" type="radio" id="pos_icono_der" value="der" checked>
	  <label for="pos_icono_der">Derecha</label>
	  <%end if%>
      </td>
    </tr>
<%end if%>

<%if config_archivo then%>
    <tr>
      <td colspan="2"><b><%=config_nom_archivo%></b></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td colspan="2">
	  <input name="archivo" id="archivo_si" type="radio" value="1"><label for="archivo_si">A&ntilde;adir</label>
	  <input name="archivo" id="archivo_no" type="radio" value="0" checked><label for="archivo_no">No a&ntilde;adir</label></td>
    </tr>
<%end if%>
  </table>
  <table width="100%"  border="0" cellspacing="0" cellpadding="2">

    <tr>
      <td align="right">&nbsp;</td>
    </tr>
    <tr>
      <td align="right">        <input name="insertar" type="submit" class="botonAdmin" id="insertar" value="Enviar"></td>
    </tr>
  </table>
  <br>
  <%if err <> 0 then Response.Write "Ha ocurrido un error.<br>"&err.description end if%>
<br>