<%' Recupero la sección (seleccionada en el listado general o al change del desplegable en esta misma zona)
	seccion = numero(request("seccion"))

	' Leo de la trabla secciones
	set re_Secciones = Server.CreateObject("ADODB.Recordset")
	re_Secciones.ActiveConnection = conn_
	if err<>0 then
		unerror = true : msgerror = "Error en conexión a base de datos.<br>"&conn_
	else
		re_Secciones.Source = "SELECT S_NOMBRE, S_ID FROM SECCIONES ORDER BY S_ORDEN"
		re_Secciones.CursorType = 3 : re_Secciones.CursorLocation = 2 : re_Secciones.LockType = 3 : re_Secciones.Open()
		if err<>0 then
			unerror = true : msgerror = "Sql:<br>" & sql &"<br><br>Error:<br>"&err.description
		end if
	end if

	' Si temenos una sección escogida, leemos y pintamos su subseccinoes.
	' Si no tiene, decimos que Escoja una.
	if seccion > 0 then
		' Insertar una sección nueva si se recibe el nombre al enviar un formulario
		nuevaseccion = ""&replace(request.Form("nuevaseccion"),"'","")
		nuevaseccion = replace(nuevaseccion,chr(34),"") ' quito las comillas dobles
		largo = len(nuevaseccion)	
		if ""&nuevaseccion <> "" then
			if largo > config_maxcarseccion then
				%><script language="javascript" type="text/javascript">alert(" * Error en longitud de nombre *\n\nEl nombre de sección escrito tiene <%=largo%> caracteres.\nPor favor escriba un nombre igual o inferior a <%=config_maxcarseccion%> caracteres.\n")</script><%
			else
				' Busco si hay otra sección con el mismo nombre
				consultaXOpen "SELECT S2_NOMBRE FROM SECCIONES2 WHERE S2_NOMBRE = '"& nuevaseccion &"' AND S2_ID_S = "& seccion,2
				if not re.eof then
					consultaXClose()
					%>
					<script language="javascript" type="text/javascript">
					f.ac.value = "adminsecciones"
					f.msgerror.value = "Hay una sub-sección con el mismo nombre para la sección indicada."
					f.submit()
					</script>
					<%
				else
					if request.Form("pos") = "inicio" then
						n = "0.9"
					else
						n = "9999"
					end if
				
					' Hago la inserción
					sql = "INSERT INTO SECCIONES2 (S2_NOMBRE, S2_ORDEN, S2_ID_S) VALUES ('"& nuevaseccion &"',"&n&","& seccion &")"
					call exeSql(sql, conn_)
					
					' Desconecto para actualizar
					set conn_activa = Nothing

					reordenaSecciones2(seccion)
					upSeccion(seccion)
					%>
					<script>
					try{
						var f1 = parent.frames[0].f // Frame de la izquierda
						f1.ac.value = ""
						f1.action = "main.asp"
						f1.target = ""
						f1.submit()
					}catch(unerror){}
					location.href='main.asp?ac=adminsecciones2&seccion=<%=seccion%>'
					</script>
					<%
				end if
			end if ' largo 
		end if
	end if ' seccion > 0

	sql = "SELECT *"
	sql = sql & " FROM SECCIONES2"
	sql = sql & " WHERE S2_ID_S = " & seccion
	sql = sql & " ORDER BY S2_ORDEN"

	consultaXOpen sql,1
	
	if not unerror then
%>
<script language="javascript" type="text/javascript">
<!--
	function changeSeccionParaSub(c){
		location.href='main.asp?ac=adminsecciones2&seccion='+c.value
	}
	function ampliarfotoseccion2(nombre){
		ventana("archivos.asp?ac=ampliarfotoseccion2&archivo="+nombre,'AmpliarFoto',100,100,0)
	}
	function ampliariconoseccion2(nombre){
		ventana("archivos.asp?ac=ampliariconoseccion2&archivo="+nombre,'AmpliarIcono',100,100,0)
	}

//-->
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF"><%=config_nom_secciones2%></font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>Escriba el nombre de la nueva secci&oacute;n y pulse en el bot&oacute;n enviar.<br>
      Recuerde que el nombre debe ser de <%=config_maxcarseccion%> car&aacute;teres como m&aacute;ximo. </td>
  </tr>
</table>
<br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<b>Sección</b>: 
	<select name="seccion" onChange="changeSeccionParaSub(this)" class="campoAdmin">
		<option value="">Seleccione ...</option>
	<%while not re_Secciones.eof%>
		<option value="<%=re_Secciones("S_ID")%>" <%if re_Secciones("S_ID") = seccion then Response.Write "selected" : seccionNombre = re_Secciones("S_NOMBRE") end if%>><%=re_Secciones("S_NOMBRE")%></option>
		<%
		re_Secciones.MoveNext()
	wend%>
	</select>
	<%
	re_Secciones.close
	set re_Secciones = Nothing
	%></td>
  </tr>
</table>
<%if seccion > 0 then%>
<br>
<span class="TituloResaltado"><%=seccionNombre%></span>
<table  border="0" cellpadding="1" cellspacing="0">
  <tr valign="middle">
    <td align="right"><input name="nuevaseccion" type="text" class="campoAdmin" id="nuevaseccion" maxlength="100">      </td>
    <td><input type="submit" class="botonAdmin" value="Enviar"></td>
    <td>
	<input name="pos" id="pos_primera" type="radio" value="inicio" <%if request.QueryString("pos") = "inicio" or ""&request.QueryString("pos") = "" then response.Write "checked" end if%>><label for="pos_primera">Primera</label>
	<input name="pos" id="pos_ultima" type="radio" value="fin" <%if request.QueryString("pos") = "fin" then response.Write "checked" end if%>><label for="pos_ultima">&Uacute;ltima</label></td>
  </tr>
</table>
<%end if%>

	<br>
	<%if request.QueryString("msgerror")&request.form("msgerror") <> "" then%>
	  <table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td align="center" class="Estilo4"><%=request.QueryString("msgerror")&request.form("msgerror")%></td>
  </tr>
</table>
	<%end if%>

	
	<%if request.QueryString("msgerror") <> "" or request.QueryString("msg") <> "" then%>
	<script>
			try{
				var f = parent.frames[0].f // Frame de la izquierda
				f.ac.value = ""
				f.action = "main.asp"
				f.target = ""
				f.submit()
			}catch(unerror){
//
			}
	</script>
	<%end if%>
	
<%if seccion <=0 then%>
	<div align="center"><b>Escoja una secci&oacute;n.</b></div>
	<%else%>
	<%if re.eof then%>
		<div align="center"><b>No hay ninguna sub-secci&oacute;n.</b></div>
		<%else%>
		
	<table width="100%"  border="0" cellspacing="2" cellpadding="2">
	  <tr class="fondoAdmin">
		<td width="13" align="center" bgcolor="#F5F5F5">&nbsp;</td>
		<td width="100%"><b>Nombre</b></td>
		<%if config_foto_seccion2 then%>
		<td align="center"><b>Foto</b></td>
		<%end if%>
		<%if config_icono_seccion2 then%>
		<td align="center"><b>Icono</b></td>
		<%end if%>
		<td align="center"><b>Registros</b></td>
		<td colspan="2" align="center" bgcolor="#f7f7f7">&nbsp;</td>
	  </tr>
	  <%while not re.eof%>
	  <tr>
		<td align="center" class="fondoOscuroAdmin" title="ID: <%=re("S2_ID")%> ALIAS: <%=re("S2_ALIAS")%> "><b><font color="#FFFFFF"><nobr><%=re("S2_ORDEN")%></nobr></font><nobr></nobr></b><nobr></nobr></td>
		<td bgcolor="#FFFFFF"><%=re("S2_NOMBRE")%></td>
		<%if config_foto_seccion2 then%>
		<td align="center" bgcolor="#FFFFFF">
			<%if re("S2_FOTO") <> "" then%>
			<a href="javascript:ampliarfotoseccion2('<%=re("S2_FOTO")%>')"><img src="img/imagen.gif" alt=" Ver foto " width="18" height="18" border="0"></a>
            <%end if%>
        </td>
		<%end if%>
		<%if config_icono_seccion2 then%>
		<td align="center" bgcolor="#FFFFFF">
			<%if re("S2_ICONO") <> "" then%>
			<a href="javascript:ampliariconoseccion2('<%=re("S2_ICONO")%>')"><img src="img/imagen.gif" alt=" Ver icono " width="18" height="18" border="0"></a>
			<%end if%>
		</td>
		<%end if%>
		<td align="center" bgcolor="#FFFFFF"><%=re("S2_REGISTROS")%></td>
		<td align="center" <%
	if re("S2_ACTIVO") then
		Response.Write "bgcolor='#95E37D' title=' Sección activada '"
	else
		Response.Write "bgcolor='#EA5940' title=' Sección desactivada '"
	end if
	%>>&nbsp;</td>
		<td align="right"><table border="0" cellspacing="0" cellpadding="0">
		  <tr>
		  <%if NOT re("S2_BLOQUEADA") then%>
			<td><a href="javascript:moverSeccion2(<%=re("S2_ID")%>,'subir')"><img src="img/flecha_arriba_h.gif" alt=" Subir " width="15" height="18" border="0"></a><a href="javascript:moverSeccion2(<%=re("S2_ID")%>,'bajar')"><img src="img/flecha_abajo_h.gif" alt=" Bajar " width="15" height="18" border="0"></a></td>
			<%end if%>
		  <%if re("S2_RENOMBRAR") then%>
			<td><a href="javascript:editarSeccion2(<%=re("S2_ID")%>,<%=seccion%>)"><img src="img/lapiz.gif" alt=" Editar " width="18" height="18" border="0"></a></td>
		  <%end if%>
		  <%if re("S2_ELIMINAR") and re("S2_REGISTROS")=0 then%>
			<td><a href="javascript:eliminarSeccion2(<%=re("S2_ID")%>,<%=seccion%>)"><img src="img/papelera.gif" alt=" Eliminar " width="18" height="18" border="0"></a></td>
		  <%end if%>
		  </tr>
		</table>      </td>
	  </tr>
	  <%re.movenext : wend%>
	</table>
	<%end if ' re.eof%>
<%end if ' seccion<= 0%>
<br>

<%	consultaXClose()
end if%>
<script>f.ac.value = "adminsecciones2"</script>