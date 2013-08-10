<!--#include virtual="/admin/usuarios/inc_gestion_grupos.asp" -->
<%
	seccionActualDelRegistro = numero(request.Form("comun"))
	seccion2ActualDelRegistro = numero(request.Form("comun2"))
	
	if cualid = "usuarios" and seccionActualDelRegistro >0 then
		set miGrupo = setGrupo(seccionActualDelRegistro)
	end if


	if request.Form("editar") <> "" then ' al pulsar insertar (o intro)
		enportada = numero(request.Form("enportada"))
		activo = numero(request.Form("activo"))
		titulo = ""&request.Form("titulo")
		seccion = ""&request.Form("seccion")
		seccion2 = ""&request.Form("seccion2")
		fuente = ""&request.Form("fuente")
		piefoto = left(replace(""&request.Form("piefoto"),"'","''"),254)
		pos_foto = request.Form("pos_foto")
		pos_icono = request.Form("pos_icono")
		cambiar_clave = bool(request.Form("cambiar_clave"))
		clave = ""&request.Form("clave")
		clave_r = ""&request.Form("clave_r")
		
		if request.Form("fecha") = 1 and ""&request.Form("fecha_dia") <> "" and ""&request.Form("fecha_mes") <> "" and ""&request.Form("fecha_ano") <> "" then
			fecha = right("0"&request.Form("fecha_dia"),2) & "/" & right("0"&request.Form("fecha_mes"),2) & "/" & right("0000"&request.Form("fecha_ano"),4)
		else
			fecha = "0:00:00"
		end if
		if ""&request.Form("hora_hora") <> "" and ""&request.Form("hora_minutos") <> "" then
			hora = Right("00"&request.Form("hora_hora"),2)& ":" &Right("00"&request.Form("hora_minutos"),2)
		else
			hora = "0:00"
		end if
		if request.Form("fechaini") = 1 and ""&request.Form("fechaini_dia") <> "" and ""&request.Form("fechaini_mes") <> "" and ""&request.Form("fechaini_ano") <> "" then
			fechaini = right("0"&request.Form("fechaini_dia"),2) & "/" & right("0"&request.Form("fechaini_mes"),2) & "/" & right("0000"&request.Form("fechaini_ano"),4)
		else
			fechaini = "0:00:00"
		end if
		
		if request.Form("fechafin") = 1 and ""&request.Form("fechafin_dia") <> "" and ""&request.Form("fechafin_mes") <> "" and ""&request.Form("fechafin_ano") <> "" then
			fechafin = right("0"&request.Form("fechafin_dia"),2) & "/" & right("0"&request.Form("fechafin_mes"),2) & "/" & right("0000"&request.Form("fechafin_ano"),4)
		else
			fechafin = "0:00:00"
		end if

		enlace = trim(request.Form("enlace"))

	
	
		sql = "UPDATE REGISTROS SET "

		' Campos fijos
		' ---------------------------------
		
		' Edito el idioma actual para las cualidad vinculadad
		if config_idioma_bd <> "" then
			sql = sql & "R_TITULO_"& session("idioma") &" = '"& replace(titulo,"'","''") &"'"
		else
			sql = sql & "R_TITULO = '"& replace(titulo,"'","''") &"'"
		end if
		
		sql = sql & ", R_SECCION = "& seccion &", R_ULTIMO_USUARIO = "& session("usuario") &", R_ULTIMA_EDICION = '"& Now() &"', R_FUENTE = '"& replace(request.Form("fuente"),"'","''") &"', R_ENLACE = '"& replace(enlace,"'","''") &"', R_PORTADA = "& enportada &", R_ACTIVO = "& activo &", R_FECHA = '"& fecha &"', R_FECHAINI = '"& fechaini &"', R_FECHAFIN = '"& fechafin &"', R_PIE_FOTO = '"& piefoto &"', R_POS_FOTO = '"& pos_foto &"', R_POS_ICONO = '"& pos_icono &"'"
		if config_activo_seccion2 and numero(seccion2) > 0 then
			sql = sql & ", R_SECCION2 = "& seccion2
		else
			sql = sql & ", R_SECCION2 = 1"
		end if
		
		if cambiar_clave then
			if clave <> "" and (clave = clave_r) then
				sql = sql & ", R_CLAVE = '"& sha256(clave) & "'"
			end if
		end if

		' Utilizo el nodo miGrupo que contiene seteado el grupo de la seccion escogida
		if cualid = "usuarios" then
			set nodosCampo = miGrupo.selectNodes("//grupos/grupo[@id="& seccion &"]//dato")
		else
			set nodosCampo = nodoCualid.childNodes
		end if

		' Campos configurables
		for each a in nodosCampo
			campo = ""&a.getAttribute("campo")
			nombrecorto = ""&a.getAttribute("nombrecorto")
			valor = ""&replace(request.Form(nombrecorto),"'","''")
			if a.nodeName = "dato" and campo <> "" then
				' Edito el idioma actual para las cualidad vinculadad
				if config_idioma_bd <> "" then
					sql = sql & ", R_"& campo &"_"& session("idioma") &"= '"& valor &"'"
				else
					sql = sql & ", R_"& campo &"= '"& valor &"'"
				end if
			end if
		next
		sql = sql & " WHERE R_ID = " & request.Form("id")

		''on error resume next
		set conn_activa = server.CreateObject("ADODB.Connection")
		conn_activa.Open conn_
		conn_activa.execute sql
		if err<>0 then
			unerror = true : msgerror = "Se ha producido un error en la SQL al intentar guardar.<br><font color=#cccccc>"&sql&"</font>"
		else

			traspasoDeSeccion seccionActualDelRegistro, seccion
			traspasoDeSeccion2 seccion2ActualDelRegistro, seccion2

			reOrdena()
			mi_seccion = seccion
			reOrdena()
			mi_seccion2 = seccion2
			reOrdena()

			' Libero conexión activa
			conn_activa.Close
			set conn_activa = nothing
			
			%>
			<p>&nbsp;</p>
			<p>Un momento ... 
			  <script language="javascript" type="text/javascript">
			try{
				var f = parent.frames[0].f // Frame de la izquierda
				f.ac.value = ""
				f.action = "main.asp"
				f.target = ""
				f.submit()

				<%select case request.Form("foto")%>
				<%case "quitar"%>
					location.href = 'archivos_frames.asp?ac=quitarfoto&id=<%=request.Form("id")%>&icono=<%=request.Form("icono")%>&archivo=<%=request.Form("archivo")%>'
				<%case "anadir"%>
					location.href = 'archivos_frames.asp?ac=formguardarfoto&id=<%=request.Form("id")%>&icono=<%=request.Form("icono")%>&archivo=<%=request.Form("archivo")%>'
				<%case "cambiar"%>
					location.href = 'archivos_frames.asp?ac=formguardarfoto&id=<%=request.Form("id")%>&icono=<%=request.Form("icono")%>&archivo=<%=request.Form("archivo")%>'
				<%case else%>

					<%select case request.Form("icono")%>
					<%case "quitar"%>
						location.href = 'archivos_frames.asp?ac=quitaricono&id=<%=request.Form("id")%>&archivo=<%=request.Form("archivo")%>'
					<%case "anadir"%>
						location.href = 'archivos_frames.asp?ac=formguardaricono&id=<%=request.Form("id")%>&archivo=<%=request.Form("archivo")%>&nombrefoto=<%=request.Form("nombrefoto")%>'
					<%case "cambiar"%>
						location.href = 'archivos_frames.asp?ac=formguardaricono&id=<%=request.Form("id")%>&archivo=<%=request.Form("archivo")%>&nombrefoto=<%=request.Form("nombrefoto")%>'
					<%case else%>
						
						<%select case request.Form("archivo")%>
						<%case "quitar"%>
							location.href = 'archivos_frames.asp?ac=quitararchivo&id=<%=request.Form("id")%>'
						<%case "anadir"%>
							location.href = 'archivos_frames.asp?ac=formguardararchivo&id=<%=request.Form("id")%>'
						<%case "cambiar"%>
							location.href = 'archivos_frames.asp?ac=formguardararchivo&id=<%=request.Form("id")%>'
						<%case "nada"%>
							location.href = 'inicio.asp'
						<%case else%>
							location.href = 'inicio.asp'
						<%end select%>
						
					<%end select%>

				<%end select%>

			}catch(unerror){
				alert(unerror.description)
			}
			  </script>
			  <%
		end if
		on error goto 0
		

	else
	
	id = request.Form("id")
	if not esNumero(id) then
		unerror = true : msgerror = "No se ha recibido una ID o no es correcta."
	end if
	
	if not unerror then
		' Abrir el registro
		sql = "SELECT * FROM REGISTROS WHERE R_ID = " & id
		consultaXOpen sql,2
		seccion = re("R_SECCION")
		if config_activo_seccion2 then
			seccion2 = re("R_SECCION2")
		end if
		

%>
              <script language="javascript" type="text/javascript">
	f.comun.value = <%=seccion%>
	f.comun2.value = <%=numero(seccion2)%>
              </script>		

		
			</p>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
		
      <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
		
      <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Editar</font></b></td>
      <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
    </tr>
	</table>
<br>
<%

	if re("R_HORA") <> 0 then
		hora = re("R_HORA")
	else
		hora = "0:00"
	end if
	
	if re("R_FECHA") <> 0 then
		fecha = re("R_FECHA")
	else
		fecha = Date()
	end if

	if ""&re("R_FECHAINI") <> "0:00:00" then
		fechaini = re("R_FECHAINI")
	else
		fechaini = Date()
	end if

	if ""&re("R_FECHAFIN") <> "0:00:00" then
		fechafin = re("R_FECHAFIN")
	else
		fechafin = Date()
	end if
	

	if config_activo_seccion2 then
		' JavaScript para las subsecciones
		Response.Write "<script language='javascript' type='text/javascript'>" & vbCrlf
		' Leo las subsecciones y las meto en un array
		set re2 = Server.CreateObject("ADODB.Recordset")
		re2.ActiveConnection = conn_
		re2.Source = "SELECT * FROM SECCIONES2 ORDER BY S2_ID_S, S2_ORDEN"
		re2.CursorType = 3 : re2.CursorLocation = 2 : re2.LockType = 3 : re2.Open()
	
			a=0
			b=0
			while not re2.eof
				if ""&estaSecc <> ""&re2("S2_ID_S") then
					b = b + 1
					arrTit = "arrTituloSeccion"& re2("S2_ID_S")
					arrVal = "arrValorSeccion"& re2("S2_ID_S")
					Response.Write "var "& arrTit &" = new Array()" & vbCrlf
					Response.Write "var "& arrVal &" = new Array()" & vbCrlf
				end if
				Response.Write arrTit&"["&a&"] = '" & re2("S2_NOMBRE") &"'"& vbCrlf
				Response.Write arrVal&"["&a&"] = '" & re2("S2_ID") &"'"& vbCrlf
				a=a+1
				estaSecc = re2("S2_ID_S")
				re2.movenext
			wend
		
		re2.Close()
		set re2 = nothing  
		Response.Write "</script>"
	end if
  %>

<script language="javascript" type="text/javascript">
	f.ac.value = "nuevo";

	function cargaSubSecciones(c){
		try{
			borrarCombo(f.seccion2);
			var arrT = eval("arrTituloSeccion"+c.value);
			var arrV = eval("arrValorSeccion"+c.value);
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


<table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td><b>Secci&oacute;n:</b>
	<%
		dim re_secc
		
		sql = "SELECT * FROM SECCIONES ORDER BY S_ORDEN"
			set re_secc = Server.CreateObject("ADODB.Recordset")
			re_secc.ActiveConnection = conn_
			if err<>0 then
				unerror = true : msgerror = "Conexion.<br>"&err.description
			else
				re_secc.Source = sql : re_secc.CursorType = 3 : re_secc.CursorLocation = 2 : re_secc.LockType = 3 : re_secc.Open()
				retotal = re_secc.recordcount
				if err<>0 then
					unerror = true : msgerror = "Sql.<br>"&err.description
				end if
			end if
		%>

<select name="seccion" class="campoAdmin" onChange="cargaSubSecciones(this)">
	<%if re_secc.recordcount > 1 then%>
        <option>Secci&oacute;n ...</option>
		<%end if%>
	  <%while not re_secc.eof
	  %>
        <option value="<%=re_secc("S_ID")%>" <%if strComp(re("R_SECCION"),re_secc("S_id"))=0 then Response.Write "selected" end if%>><%=re_secc("S_NOMBRE")%></option>
		<%re_secc.movenext : wend%>
      </select>
	
		<%
		re_secc.Close() : set re_secc = Nothing
		
		if config_activo_seccion2 then
			' Cargao las subseccines
			sql = "SELECT * FROM SECCIONES2 WHERE S2_ID_S = "& seccion &" ORDER BY S2_ORDEN"
			set re_sub = Server.CreateObject("ADODB.Recordset")
			re_sub.ActiveConnection = conn_
			re_sub.Source = sql : re_sub.CursorType = 3 : re_sub.CursorLocation = 2 : re_sub.LockType = 3 : re_sub.Open()
			
			if not re_sub.eof then%>
			/
			<select name="seccion2" class="campoAdmin">
			<option value="1">Sub secci&oacute;n ...</option>
			<%while not re_sub.eof%>
			<option value="<%=re_sub("S2_ID")%>" <%if strComp(seccion2,re_sub("S2_ID"))=0 then Response.Write "selected" end if%>><%=re_sub("S2_NOMBRE")%></option>
			<%re_sub.movenext : wend%>
			</select>
			<%end if
			re_sub.Close() : set re_sub = Nothing
		end if%>
		<input name="seccion_actual" type="hidden" value="<%=request.Form("seccion")%>">		
		<input name="seccion_actual2" type="hidden" value="<%=request.Form("seccion2")%>">		
		</td>

    <td align="right">
<%if config_portada then%>
	<table  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><input name="enportada" type="checkbox" id="enportada" value="1" <%if re("R_PORTADA") then Response.Write "checked" end if%>></td>
        <td><img src="img/bandera.gif" width="18" height="18"></td>
        <td><label for="enportada"><b>En portada </b></label></td>
      </tr>
    </table>
<%end if%>
	
<%if config_activo then%>
<script language="javascript" type="text/javascript">
	function clickActivo(a){
//		try{
//			f.fechaini.click()
//		}catch(unerror){
//			alert(unerror)
//		}
	}
</script>
<table  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><input name="activo" type="checkbox" id="activo" onClick="clickActivo(this.checked)" value="1" <%if re("R_ACTIVO") then Response.Write "checked" end if%>></td>
        <td><label for="activo"><b>Activo</b></label></td>
      </tr>
    </table>
<%end if%>
</td>
	
  </tr>
  
  
  	<%if config_fechainifin then%>
<tr bgcolor="#FFFFFF">
      <td colspan="2">&nbsp;<b><%=config_nom_fechaini%></b>
	 <input name="fechaini" id="fechaini_si" type="radio" onClick="siAnadirFechaini()" value="1" <%if ""&re("R_FECHAINI") <> "0:00:00" then Response.Write "checked" end if%>><label for="fechaini_si">S&iacute;&nbsp;</label>
	 <select name="fechaini_dia" <%if ""&re("R_FECHAINI") = "0:00:00" then Response.Write "disabled='disabled'" end if%>  class="campoAdmin" id="fechaini_dia">
          <% for n=1 to 31 %>
          <option value="<%=n%>" <%if n = day(fechaini) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
        <select name="fechaini_mes" <%if ""&re("R_FECHAINI") = "0:00:00" then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechaini_mes">
          <% for n=1 to 12 %>
          <option value="<%=n%>"<%if n = month(fechaini) then Response.Write "selected" end if%>><%=arrMes(n-1)%></option>
          <% next %>
        </select>
        <select name="fechaini_ano" <%if ""& re("R_FECHAINI") ="0:00:00" then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechaini_ano">
          <% for n=Year(date())-12 to Year(date())+1%>
          <option value="<%=n%>"<%if n = year(fechaini) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
	  <input name="fechaini" id="fechaini_no" type="radio" onClick="noAnadirFechaini()" value="0" <%if ""&re("R_FECHAINI") = "0:00:00" or ""&fechaini = "" then Response.Write "checked" end if%>><label for="fechaini_no">No</label></td>
  </tr>
	
<tr bgcolor="#FFFFFF">
      <td colspan="2">
	    <b>&nbsp;<%=config_nom_fechafin%></b>
	    <input name="fechafin" id="fechafin_si" type="radio" onClick="siAnadirFechafin()" value="1" <%if ""&re("R_FECHAFIN") <> "0:00:00" then Response.Write "checked" end if%>><label for="fechafin_si">S&iacute;&nbsp;</label>
		<select name="fechafin_dia" <%if ""& re("R_FECHAFIN") ="0:00:00" then Response.Write "disabled='disabled'" end if%>  class="campoAdmin" id="fechafin_dia">
          <% for n=1 to 31 %>
          <option value="<%=n%>" <%if n = day(fechafin) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
        <select name="fechafin_mes" <%if ""& re("R_FECHAFIN") ="0:00:00" then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechafin_mes">
          <% for n=1 to 12 %>
          <option value="<%=n%>"<%if n = month(fechafin) then Response.Write "selected" end if%>><%=arrMes(n-1)%></option>
          <% next %>
        </select>
        <select name="fechafin_ano" <%if ""& re("R_FECHAFIN")="0:00:00" then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fechafin_ano">
          <% for n=Year(date())-12 to Year(date())+1%>
          <option value="<%=n%>"<%if n = year(fechafin) then Response.Write "selected" end if%>><%=n%></option>
          <% next %>
        </select>
	  <input name="fechafin" id="fechafin_no" type="radio" onClick="noAnadirFechafin()" value="0" <%if ""&re("R_FECHAFIN") = "0:00:00" then Response.Write "checked" end if%>><label for="fechafin_no">No</label></td>
  </tr>
	
	
	<%end if%>	

  
  <%if config_fecha then%>
  <tr>
    <td colspan="2"><b>&nbsp;<%=config_nom_fecha%></b>
	<%if  config_fecha<>2 then%>
      <input name="fecha" id="fecha_si" type="radio" onClick="siAnadirFecha()" value="1" <%if re("R_FECHA") <> 0  then Response.Write "checked" end if%>><label for="fecha_si">S&iacute;&nbsp;</label>
	  <%else%>
	  <input type="hidden" name="fecha" value="1">
	  <%end if%>
	  <select name="fecha_dia" <%if not re("R_FECHA") <> 0 and config_fecha<>2 then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fecha_dia">
        <% for n=1 to 31 %>
        <option value="<%=n%>" <%if n = day(fecha) then Response.Write "selected" end if%>><%=n%></option>
        <% next %>
      </select>
      <select name="fecha_mes" <%if not re("R_FECHA") <> 0 and config_fecha<>2 then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fecha_mes">
        <% for n=1 to 12 %>
        <option value="<%=n%>"<%if n = month(fecha) then Response.Write "selected" end if%>><%=arrMes(n-1)%></option>
        <% next %>
      </select>
      <select name="fecha_ano" <%if not re("R_FECHA") <> 0 and config_fecha<>2 then Response.Write "disabled='disabled'" end if%> class="campoAdmin" id="fecha_ano">
        <% for n=Year(date())-12 to Year(date())+1%>
        <option value="<%=n%>"<%if n = year(fecha) then Response.Write "selected" end if%>><%=n%></option>
        <% next %>
      </select>
	  <%if  config_fecha<>2 then%>
      <input name="fecha" id="fecha_no" type="radio" onClick="noAnadirFecha()" value="0" <%if not re("R_FECHA") <> 0 then Response.Write "checked" end if%>><label for="fecha_no">No</label>
	  <%end if%></td>
  </tr>
  <%end if%>
 <%if config_hora then%>
  <tr>
    <td colspan="2"><b>&nbsp;Hora </b>
	    <select name="hora_hora" class="campoAdmin" id="hora_hora">
          <% for n=0 to 23 %>
          <option value="<%=n%>" <%if n=HOUR(hora) then  Response.Write "selected" end if%>   ><%=Right("00"&n,2)%></option>
          <% next %>
        </select>
	    <strong>: </strong><select name="hora_minutos" class="campoAdmin" id="hora_minutos">
          <% for n=0 to 11 %>
          <option value="<%=n*5%>" <%if n*5=MINUTE(hora) then  Response.Write "selected" end if%> ><%=Right("00"&n*5,2)%></option>
          <% next %>
        </select>
  </tr>
  <%end if%>
  <tr>
    <td colspan="2"><b><%=config_nom_titulo%></b>* </td>
  </tr>

<%if cualid = "usuarios" then%>
	<tr bgcolor="#FFFFFF">
	<td colspan="2"><span title="No es posible cambiar el nombre de usuario."><%=re("R_TITULO")%></span>
    <input name="titulo" type="hidden" id="titulo" value="<%=re("R_TITULO")%>">
	</td>
  </tr>
<%else%>
	<tr>
    <td colspan="2">
	<%
	if ""&config_idioma_bd = "" then
		titulo = re("R_TITULO")
	else
		titulo = re("R_TITULO_"& session("idioma"))
	end if
	%>
    <input name="titulo" type="text" class="campoAdmin" id="titulo" style="width:100%" value="<%=titulo%>" maxlength="255">
	</td>
	</tr>
<%end if%>
  
  <%if cualid = "usuarios" then%>
  <tr>
    <td colspan="2"><table width="100%"  border="0" cellspacing="0" cellpadding="2">
      <tr>
        <td><table  border="0" align="center" cellpadding="4" cellspacing="0" class="fondoAdmin">
          <tr>
            <td><table  border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td>
				  <script language="javascript" type="text/javascript">
				  function cambiarClave(c){
				  	f.clave.disabled = !c.checked
				  	f.clave_r.disabled = !c.checked
					if (c.checked) f.clave.focus()
				  }
				  </script>
				  <input name="cambiar_clave" type="checkbox" id="cambiar_clave" value="1" onClick="cambiarClave(this)"></td>
                  <td align="left"><label for="cambiar_clave"><nobr>Cambiar clave&nbsp;</nobr></label></td>
                </tr>
            </table></td>
            <td align="left"><b>Clave</b>
                <input name="clave" type="password" disabled="true" class="campoAdmin" id="clave" size="10" maxlength="255">
&nbsp;<nobr><b>Repetir clave*</b></nobr>
      <input name="clave_r" type="password" disabled="true" class="campoAdmin" id="clave_r" size="10" maxlength="255">
&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table>
    </td>
  </tr>

<%if re("R_SECCION") = 2 THEN%>
  <tr>
    <td colspan="2"><b>Permisos</b></td>
  </tr>
  <tr>
    <td colspan="2">
			<button onClick="ventana('usuarios_personalizados.asp?ac=editar&id=<%=re("R_ID")%>','UsuariosPersonalizados',475,575,1);" type="button" name="" style="width:120px;cursor:default">
			<table border="0" cellpadding="1" cellspacing="0" width="100%">
				<tr valign="middle">
				<td align="left"><img src="../images/candado.gif" width="12" height="15" hspace="2"></td>
				<td><nobr>Editar permisos</nobr></td>
				</tr>
			</table>
			</button>
	</td>
  </tr>
  <%end if%>

  <%end if%>
  

  
  <!-- Campo configurables -->
<%

	' Consulta para contenido del campo ORDENIDIOMA (si esta activado el parametro)
	if config_ordenidioma then
		sql = "SELECT * FROM ORDENIDIOMA"
		set re_oi = Server.CreateObject("ADODB.Recordset")
		re_oi.ActiveConnection = conn_
		re_oi.Source = sql : re_oi.CursorType = 3 : re_oi.CursorLocation = 2 : re_oi.LockType = 3 : re_oi.Open()
	end if

	' Utilizo el nodo miGrupo que contiene seteado el grupo de la seccion escogida
	if cualid = "usuarios" then
		set nodosCampo = miGrupo.selectNodes("//grupos/grupo[@id="& seccion &"]//dato")
	else
		set nodosCampo = nodoCualid.childNodes
	end if

'on error resume next
numCampo = 0
for each a in nodosCampo
	c_titulo = ""&a.getAttribute("titulo")
	nombrecorto = ""&a.getAttribute("nombrecorto")
	campo = ""&a.getAttribute("campo")
	c_tipo = ""&a.getAttribute("tipo")
	numCampo = numCampo + 1
	if a.nodeName = "dato" and campo <> "" then
%>

  <tr>
    <td colspan="2"><b><%=c_titulo%></b><%if ""&a.getAttribute("requerido")="1" then%>*<%end if%></td>
  </tr>
  <tr>
    <td colspan="2" valign="top">
		<%
		if config_idioma_bd = "" then
			r_campo = re("R_" & campo)
		else
			r_campo = re("R_" & campo & "_" & session("idioma"))
		end if

	select case c_tipo
	  case "texto"%>
      <input name="<%=nombrecorto%>" type="text" <%if a.getAttribute("manipulable")="0" then response.Write("disabled='true'") end if%> class="campoAdmin" id="subtitulo" style="width:100%" value="<%=r_campo%>" maxlength="255">
      <%case "memo"%>
      <textarea name="<%=nombrecorto%>" rows="<%=a.getAttribute("filas")%>" wrap="virtual" class="areaAdmin" style="width:100%"><%=r_campo%></textarea>
	<%if a.getAttribute("editorhtml") = 1 then%>
	  <script language="javascript1.2">
		if (document.getElementById("<%=nombrecorto%>")) {
			editor_generate("<%=nombrecorto%>");
		}
	  </script>
	  <%end if%>
	  <%case "combo"%>
	  <select name="<%=nombrecorto%>" class="campoAdmin">
	  <option value="">Seleccione una ...</option>
	  <%for each opcion in a.childNodes%>
	  	<option value="<%=opcion.getAttribute("valor")%>" <%if r_campo = opcion.getAttribute("valor") then Response.Write "selected" end if%>><%=opcion.getAttribute("titulo")%></option>
	  <%next%>
	  </select>

	  <%case "opcion"%>
	  	<%
		inc = 0
		for each opcion in a.childNodes
		inc = inc + 1%>
		<input type="radio" name="<%=nombrecorto%>" id="<%=campo&inc%>" value="<%=opcion.getAttribute("valor")%>" <%if r_campo = opcion.getAttribute("valor") then Response.Write "checked" end if%>><label for="<%=campo&inc%>"><%=opcion.getAttribute("titulo")%></label>
		<%next%>

	  <%case "check"
	  	valores = split(""&r_campo,",")
		for each valor in valores
			cadena = cadena &"|"& trim(valor)
		next
		cadena = cadena &"|"
		inc = 0
	  	for each opcion in a.childNodes
		inc = inc + 1%>
		<input  type="checkbox" name="<%=nombrecorto%>" id="<%=campo&inc%>" value="<%=opcion.getAttribute("valor")%>" <%if inStr(cadena,opcion.getAttribute("valor"))>0 then Response.Write "checked" end if%>><label for="<%=campo&inc%>"><%=opcion.getAttribute("titulo")%></label>
		<%next%>

	<%case "color"%>
		<div align="center">
		<input name="<%=nombrecorto%>" type="text" class="campoAdmin" value="<%=r_campo%>" size="8" maxlength="7" readonly="true">
		<div id="capaColorFlash" style="position:absolute; z-index:<%=10-numCampo%>;width: 172; height: 155; visibility: visible;">
		<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="172" height="155">
		<param name="movie" value="color.swf?nombre=<%=nombrecorto%>&col=<%=Replace(re("R_" & campo),"#","")%>">
		<param name="quality" value="high">
		<param name="WMODE" value="transparent">
		<embed src="color.swf?nombre=<%=nombrecorto%>&col=<%=Replace(r_campo,"#","")%>" width="172" height="155" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent"></embed>
		</object>
		</div>
		<br>
		</div>

	<%case "orden_idioma"

	if config_ordenidioma then%>
	
	<select name="orden_idioma" class="campoAdmin">
		<option value="">Seleccione una ...</option>
		<% while not re_oi.eof%>
		<option value="<%=re_oi("OI_TITULO")%>" <%if re("R_ORDEN_IDIOMA") = re_oi("OI_TITULO") then Response.Write "selected" end if%>><%=re_oi("OI_TITULO")%></option>
		<%re_oi.movenext : wend%>
	</select>

<%	end if

case else
	if c_tipo = "" then%>
		<font color="#FF0000">No se ha difinido un tipo para este campo.</font>
	<%else%>
		<font color="#FF0000">El tipo de campo <%=c_tipo%> no es v&aacute;lido.</font>
	<%end if
end select%>
    </td>
  </tr>
  <%end if
  if err<>0 then
  	%>
	<tr><td><font color="#FF0000" size="1">^^ Error en este campo.</font><a href="#DescError">ver error</a></td>
	</tr>
	<%
  	unerror = true : msgerror = err.description & "<br>Campo: "& campo & "<br>Nombre corto: " & nombrecorto
  	exit for
  end if
 next
   on error goto 0
  
  %>

  <!-- FIN: Campo configurables -->
<%if config_fuente then%>
  <tr>
    <td colspan="2" valign="top"><b><%=config_nom_fuente%></b></td>
  </tr>
  <tr>
    <td colspan="2" valign="top"><input name="fuente" type="text" class="campoAdmin" id="fuente" style="width:100%" value="<%=re("R_FUENTE")%>" maxlength="255"></td>
  </tr>
  <tr>
    <td colspan="2" valign="top"><b><%=config_nom_enlace%></b></td>
  </tr>
  <tr>
    <td colspan="2" valign="top"><input name="enlace" type="text" class="campoAdmin" id="enlace" style="width:100%" value="<%=re("R_ENLACE")%>" maxlength="255"></td>
  </tr>
  <%end if%>
<%if config_foto then%>
    <tr>
      <td colspan="2"><b><%=config_nom_foto%></b></td>
    </tr>
	<%if ""&re("R_FOTO") <> "" then%>
    <tr bgcolor="#FFFFFF" >
      <td><input type="hidden" name="nombrefoto" value="<%=re("R_FOTO")%>"><a href="javascript:ampliarfoto('<%=re("R_FOTO")%>')"><img src="img/imagen.gif" alt=" Ver foto ampliada. " width="18" height="18" border="0" align="absmiddle"></a>
        <input name="foto" id="foto_cambiar" type="radio" value="cambiar"><label for="foto_cambiar">Cambiar</label>
		<input name="foto" id="foto_quitar" type="radio" value="quitar"><label for="foto_quitar">Quitar</label>
        <input name="foto" id="foto_nada" type="radio" value="nada" checked ><label for="foto_nada">Nada</label></td>
      
	  
	  <td><%if config_posicion_foto then%>
	  <input name="pos_foto" id="pos_foto_izq" type="radio" value="izq" <%if re("R_POS_FOTO") = "izq" then%>checked<%end if%>>
          <label for="pos_foto_izq">Izquierda</label>
          <input name="pos_foto" type="radio" id="pos_foto_der" value="der" <%if re("R_POS_FOTO") = "der" then%>checked<%end if%>>
          <label for="pos_foto_der">Derecha</label>
		  <%end if%></td>
	 
		  
    </tr>
	<%else%>
    <tr bgcolor="#FFFFFF">
      <td>
	  <input name="foto" id="foto_anadir" type="radio" value="anadir"><label for="foto_anadir">A&ntilde;adir</label>
	  <input name="foto" id="foto_nada" type="radio" value="nada" checked ><label for="foto_nada">No a&ntilde;adir</label></td>
      
	  <td>
		<%if config_posicion_foto then%>
        <input name="pos_foto" id="pos_foto_izq" type="radio" value="izq" <%if re("R_POS_FOTO") = "izq" then%>checked<%end if%>>
        <label for="pos_foto_izq">Izquierda</label>
		<input name="pos_foto" type="radio" id="pos_foto_der" value="der" <%if re("R_POS_FOTO") = "der" then%>checked<%end if%>>
        <label for="pos_foto_der">Derecha</label>
		<%end if%>
	  </td>
		
		
    </tr>

  <%end if%>
	<%if config_pie_foto then%>
    <tr>
      <td colspan="2"><b>Pie de foto:</b></td>
    </tr>
    <tr>
      <td colspan="2"><textarea name="piefoto" cols="" rows="3" wrap="virtual" class="areaAdmin" id="piefoto" style="width:100%"><%=re("R_PIE_FOTO")%></textarea></td>
    </tr>
	<%end if%>
  <%end if%>
  
  <%if config_icono then%>
    <tr>
      <td colspan="2"><b><%=config_nom_icono%></b></td>
    </tr>
	<%if ""&re("R_icono") <> "" then%>
    <tr bgcolor="#FFFFFF">
      <td><a href="javascript:ampliaricono('<%=re("R_icono")%>')"><img src="img/imagen.gif" alt=" Ver icono ampliado. " width="18" height="18" border="0" align="absmiddle"></a>
        <input name="icono" type="radio" id="icono_auto" value="auto">
        <label for="icono_auto">Auto</label>
        <input name="icono" id="icono_cambiar" type="radio" value="cambiar"><label for="icono_cambiar">Cambiar</label>
		<input name="icono" id="icono_quitar" type="radio" value="quitar"><label for="icono_quitar">Quitar</label>
	  <input name="icono" id="icono_nada" type="radio" value="nada" checked ><label for="icono_nada">Nada</label></td>

      
	  <td>
	  <%if config_posicion_icono then%>
	  <input name="pos_icono" type="radio" id="pos_icono_izq" value="izq" <%if re("R_POS_ICONO") = "izq" then%>checked<%end if%>>
          <label for="pos_icono_izq">Izquierda</label>
          <input name="pos_icono" type="radio" id="pos_icono_der" value="der" <%if re("R_POS_ICONO") = "der" then%>checked<%end if%>>
          <label for="pos_icono_der">Derecha</label>
		  <%end if%>
	  </td>
		 
    </tr>
	<%else%>
    <tr bgcolor="#FFFFFF">
      <td>
	  <%if config_iexplorer then%>
	  <input name="icono" type="radio" id="icono_auto" value="auto">
	  <label for="icono_auto">Auto</label>
	  <%end if%>
	  <input name="icono" id="icono_anadir" type="radio" value="anadir"><label for="icono_anadir">A&ntilde;adir</label>
	  <input name="icono" type="radio" id="icono_nada" value="nada" checked >
	  <label for="icono_nada">No a&ntilde;adir</label></td>
      
	  
	  <td>
	  <%if config_posicion_icono then%>
	  <input name="pos_icono" type="radio" id="pos_icono_izq" value="izq" <%if re("R_POS_ICONO") = "izq" then%>checked<%end if%>>
	  <label for="pos_icono_izq">Izquierda</label>
	  <input name="pos_icono" type="radio" id="pos_icono_der" value="der" <%if re("R_POS_ICONO") = "der" then%>checked<%end if%>>
	  <label for="pos_icono_der">Derecha</label>
	  <%end if%>
	  </td>
	  
	  
    </tr>
  <%end if%>

  <%end if%>

  <%if config_archivo then%>
    <tr>
      <td colspan="2"><b><%=config_nom_archivo%></b></td>
    </tr>
	<%if ""&re("R_archivo") <> "" then%>
    <tr bgcolor="#FFFFFF">
      <td><%pintaIconoExtension(re("R_TIPOARCHIVO"))%>
        <input name="archivo" id="archivo_cambiar" type="radio" value="cambiar"><label for="archivo_cambiar">Cambiar</label>
		<input name="archivo" id="archivo_quitar" type="radio" value="quitar"><label for="archivo_quitar">Quitar</label>
		<input name="archivo" id="archivo_nada" type="radio" value="nada" checked ><label for="archivo_nada">Nada</label></td>
      <td>&nbsp;</td>
    </tr>
	<%else%>
    <tr bgcolor="#FFFFFF">
      <td>
	  <input name="archivo" id="archivo_anadir" type="radio" value="anadir"><label for="archivo_anadir">A&ntilde;adir</label>
	  <input name="archivo" id="archivo_nada" type="radio" value="nada" checked ><label for="archivo_nada">No a&ntilde;adir</label></td>
      <td>&nbsp;</td>
    </tr>
  <%end if%>

  <%end if%>
 
  <tr>
    <td colspan="2" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="right" valign="top"><input name="" type="button" class="botonAdmin" onClick="location.href='inicio.asp'" value="Cancelar">
    <input name="editar" type="submit" disabled="true" class="botonAdmin" id="editar" value="Enviar">
	</td>
  </tr>
</table>
<script>
function cargada(){
	f.editar.disabled = false
}
f.ac.value = "editar"
f.id.value = "<%=id%>"
window.onload = cargada;
</script>
<br>

		<%
	end if



	if not unerror then
		consultaXClose()
	end if
	
	end if ' al pulsar insertar (o intro)

	if unerror then
%>

<a name="DescError">
        <table width="400"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#FF0000">
          <tr>
            <td><table width="100%"  border="0" cellpadding="8" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td align="left"><font color="#666666"><b>ATENCI&Oacute;N</b></font><br>                  <%=msgerror%></td>
              </tr>
            </table></td>
          </tr>
        </table>
        <br>
		<br>

<%end if%>