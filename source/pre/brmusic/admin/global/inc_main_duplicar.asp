
	<script>
	f.ac.value = "duplicar"
	f.id.value = <%=numero(request.Form("id"))%>
	</script>
	
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
	<td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Duplicar registro</font></b></td>
	<td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
	</tr>
</table>
<br>
<%
	comp_cualid = ""&request.Form("comp_cualid")
	comp_seccion = numero(request.Form("comp_seccion"))
	comp_seccion2 = numero(request.Form("comp_seccion2"))
	id = numero(request.Form("id"))

	' Cadena de conexión
	if ""&comp_cualid <> "" then
		conn = replace(conn_,cualid,comp_cualid)
	else
		comp_cualid = cualid
		conn = conn_
	end if

	
	if lcase(request.Form("enviar")) = "enviar" then
		set origen = Server.CreateObject("ADODB.Recordset")
		origen.ActiveConnection = conn_
		origen.Source = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
		origen.CursorType = 3 : origen.CursorLocation = 2 : origen.LockType = 2
		origen.Open()

		if origen.eof then
			unerror = true : msgerror = "No se ha encontrado el registro indicado."
		else

			for each campo in nodoCualid.selectNodes("campo")
				nombre = "R_"& ucase(campo.getAttribute("nombre"))
				resto_nombres = resto_nombres &", "& nombre
				resto_valores = resto_valores &", '"& replace(origen(nombre),"'","''") &"'"
			next

			if origen("R_PORTADA") then
				portada = 1
			else
				portada = 0
			end if
			if origen("R_ACTIVO") then
				activo = 1
			else
				activo = 0
			end if

			insertar = insertarRegistro(origen("R_TITULO"), request.Form("comp_seccion"), request.Form("comp_seccion2"), session("usuario"),origen("R_FUENTE"), 1, portada, activo, origen("R_ENLACE"), origen("R_FECHA"), origen("R_FECHAINI"), origen("R_FECHAFIN"), resto_nombres, resto_valores, conn)
			if insertar <> "" then
				unerror = true : msgerror = "No se ha logrado insertar la copia de su registro en la nueva ubicación.<br>" & insertar
			end if

			' Duplicar las fotos si dispone.
			foto = ""& origen("R_FOTO")
			icono = ""& origen("R_ICONO")
			tipoFoto = getExtension(foto)
			tipoIcono = getExtension(icono)

			' Seleciono el ultimo registro (el de ahroa)
			if not unerror then
				if foto <> "" or icono <> "" then
					set destino = Server.CreateObject("ADODB.Recordset")
					destino.ActiveConnection = conn
					destino.Source = "SELECT TOP 1 R_ID, R_FOTO, R_ICONO FROM REGISTROS ORDER BY R_ID DESC"
					destino.CursorType = 3 : destino.CursorLocation = 2 : destino.LockType = 2
					destino.Open()
					if not destino.eof then
						id = destino("R_ID")
						if foto <> "" then
							destino("R_FOTO") = "foto"& id &"."& tipoFoto
						end if
						if icono <> "" then
							destino("R_ICONO") =  "icono"& id &"."& tipoIcono
						end if
						destino.Update()
					end if
					destino.Close()			
					set destino = nothing
				end if
			end if
			
			if foto <> "" or icono <> "" then
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				if foto <> "" then
					tipo = getExtension(foto)
					if fso.FileExists(Server.MapPath("/"& c_s &"datos/"& session("idioma") &"/"& cualid &"/fotos/"& foto)) then
						set fileFoto = fso.getFile(Server.MapPath("/"& c_s &"datos/"& session("idioma") &"/"& cualid &"/fotos/"& foto))
						fileFoto.Copy Server.MapPath("/"& c_s &"datos/"& session("idioma") &"/"& comp_cualid &"/fotos/foto"& id &"."& tipo)
					end if
				end if
				if icono <> "" then
					tipo = getExtension(icono)
					if fso.FileExists(Server.MapPath("/"& c_s &"datos/"& session("idioma") &"/"& cualid &"/iconos/"& icono)) then
						set fileFoto = fso.getFile(Server.MapPath("/"& c_s &"datos/"& session("idioma") &"/"& cualid &"/iconos/"& icono))
						fileFoto.Copy Server.MapPath("/"& c_s &"datos/"& session("idioma") &"/"& comp_cualid &"/iconos/icono"& id &"."& tipo)
					end if
				end if
				set fso = Nothing
			end if

		end if ' eof

		origen.Close()
		set origen = Nothing
		
		if not unerror then%>
			Un momento ...
			<script>
			try{
				var f = parent.frames[0].f // Frame de la izquierda
				f.ac.value = ""
				f.action = "main.asp"
				f.target = ""
				f.submit()
				location.href = 'inicio.asp'
			}catch(unerror){alert(unerror)}
			</script>
		<%else
			Response.Write "<b>Error:</b> " & msgerror
		end if
		
	else

		' Nodo de compatibles
		set nodoCompatibles = nodoCualid.selectSingleNode("compatible")
		if typeOK(nodoCompatibles) then
			num_compatibles = nodoCompatibles.childNodes.length
		else
			num_compatibles = 0
		end if
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="2">
	<tr>
	<td width="150">&nbsp;</td>
	<td>Seleccione ... </td>
  </tr>
  <tr>
    <td align="right">

	
	<b>Zona:&nbsp;</b></td>
    <td>
	<%if num_compatibles > 0 then%>
	<select name="comp_cualid" class="campoAdmin" onChange="submit()">
		<option value="<%=cualid%>" <%if comp_cualid = cualid then Response.Write "selected" end if%>><%=ucase(cualid)%> (Actual)</option>
		<%for each cCualid in nodoCompatibles.childNodes%>
		<option value="<%=cCualid.nodeName%>" <%if request.Form("comp_cualid") = cCualid.nodeName then Response.Write "selected" end if%>><%=ucase(cCualid.nodeName)%></option>
		<%next%>
	  </select>
	<%else
	  comp_cualid = cualid%>
	<b><%=ucase(cualid)%></b>
	<%end if%>
	</td>
  </tr>


  <%if comp_cualid <> "" then

  	' Leer datos de la nueva
	set nodoCompCualid = cualidades.selectsingleNode(comp_cualid)
	if not typeOK(nodoCompCualid) then
		unerror = true : msgerror = "La zona seleccionada no está disponible."
	else
		set nodoSecciones = nodoCompCualid.selectSingleNode("secciones")
		if not typeOK(nodoSecciones) then
			activo_seccion2 = false
		else
			activo_seccion2 = cbool(""&nodoSecciones.getAttribute("activo"))
		end if
	end if
	
	if not unerror then%>
		<tr>
		<td align="right"><b>Secci&oacute;n:&nbsp;</b></td>
		<td>
		<%
		nConn_ = replace(conn_,cualid,comp_cualid)
		on error resume next
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = nConn_
		if err <> 0 then
			unerror = true : msgerror = "No se ha encontrado la Base de datos para la Zona indicada."
		end if
		
		if not unerror then
			re.Source = "SELECT * FROM SECCIONES ORDER BY S_ORDEN"
			re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 2
			re.Open()
			if re.eof then%>
				<b>No hay secciones</b>
			<%else%>
				<select name="comp_seccion" class="campoAdmin" onChange="submit();">
						<option value="1">Sección ...</option>
					<%while not re.eof%>
						<option value="<%=re("S_ID")%>" <%if comp_seccion = re("S_ID") then Response.Write "selected" end if%>><%=re("S_NOMBRE")%></option>
						<%re.movenext()
					wend%>
				</select>
			<%end if
			re.Close()
		end if
		set re = nothing
	  %>
	  </td>
	  </tr>
	<%end if%>
  <%end if%>
  
  <%if comp_cualid <> "" and comp_seccion > 0 then%>
  <tr>
    <td align="right"><b>Sub secci&oacute;n:&nbsp;</b></td>
    <td>
	<%
	set re = Server.CreateObject("ADODB.Recordset")
	re.ActiveConnection = nConn_
	re.Source = "SELECT * FROM SECCIONES2 WHERE S2_ID_S = "& comp_seccion &" ORDER BY S2_ORDEN"
	re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 2
	re.Open()
	if re.eof then%>
		<b>No hay secciones</b>
	<%else%>
		<select name="comp_seccion2" class="campoAdmin" onChange="submit();">
			<option value="1">Subsección ...</option>
			<%while not re.eof%>
				<option value="<%=re("S2_ID")%>" <%if comp_seccion2 = re("S2_ID") then Response.Write "selected" end if%>><%=re("S2_NOMBRE")%></option>
				<%re.movenext()
			wend%>
		</select>
	<%end if
	re.Close()
	set re = nothing
  %>
  </td>
  </tr>
  <%end if%>
  
  <tr>
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>

<%	' Tengo Cualidad y sección y no tengo acceso a secciones2
if comp_cualid <> "" and comp_seccion > 0 and not activo_seccion2 then%>
		<tr>
		<td align="right">&nbsp;</td>
		<td><input name="enviar" type="submit" class="botonAdmin" value="Enviar"></td>
		</tr>
<%' Tengo todo ...
elseif comp_cualid <> "" and comp_seccion > 0 and comp_seccion2 > 1 then%>
		<tr>
		<td align="right">&nbsp;</td>
		<td><input name="enviar" type="submit" class="botonAdmin" value="Enviar"></td>
		</tr>
<%end if%>

  
</table>

<%end if%>
<br>