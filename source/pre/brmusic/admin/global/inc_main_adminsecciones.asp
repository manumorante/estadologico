<%' Insertar una sección nueva si se recibe el nombre al enviar un formulario
	seccion = ""&replace(request.Form("nuevaseccion"),"'","")
	seccion = replace(seccion,chr(34),"") ' quito las comillas dobles
	largo = len(seccion)
	if ""&seccion <> "" then
		if largo > config_maxcarseccion then
			%><script language="javascript" type="text/javascript">alert(" * Error en longitud de nombre *\n\nEl nombre de sección escrito tiene <%=largo%> caracteres.\nPor favor escriba un nombre igual o inferior a <%=config_maxcarseccion%> caracteres.\n")</script><%
		else
			' Busco si hay otra sección con el mismo nombre
			consultaXOpen "SELECT S_NOMBRE FROM SECCIONES WHERE S_NOMBRE = '"& seccion &"'",2
			if not re.eof then
				consultaXClose()
				%>
				<script>
				f.ac.value = "adminsecciones"
				f.msgerror.value = "Hay una sección con el mismo nombre."
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
				sql = "INSERT INTO SECCIONES (S_NOMBRE, S_ORDEN, S_NUEVOS) VALUES ('"& seccion &"',"&n&",1)"
				call exeSql(sql, conn_)
				set conn_activa = nothing
				reordenaSecciones()
				%>
				<script language="javascript" type="text/javascript">
				try{
					var f1 = parent.frames[0].f // Frame de la izquierda
					f1.ac.value = ""
					f1.action = "main.asp"
					f1.target = ""
					f1.submit()
				}catch(unerror){}
				f.ac.value = "adminsecciones"
				f.submit()
				</script>
				<%
			end if
		end if ' largo 
	end if
	

	sql = "SELECT *"
	sql = sql & " FROM SECCIONES"
	sql = sql & " ORDER BY S_ORDEN"

	consultaXOpen sql,1
	
	if not unerror then
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF"><%=config_nom_secciones%></font></b></td>
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
<table  border="0" cellpadding="1" cellspacing="0">
  <tr valign="middle">
    <td align="right"><input name="nuevaseccion" type="text" class="campoAdmin" id="nuevaseccion" maxlength="100">      </td>
    <td><input type="submit" class="botonAdmin" value="Enviar"></td>
    <td>
	<input name="pos" id="pos_primera" type="radio" value="inicio" <%if request.QueryString("pos") = "inicio" or ""&request.QueryString("pos") = "" then response.Write "checked" end if%>><label for="pos_primera">Primera</label>
	<input name="pos" id="pos_ultima" type="radio" value="fin" <%if request.QueryString("pos") = "fin" then response.Write "checked" end if%>><label for="pos_ultima">&Uacute;ltima</label></td>
  </tr>
</table>
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
	
<%if re.eof then%>
	<div align="center"><b>No hay ninguna secci&oacute;n.</b></div>
<%else%>
	<script language="javascript" type="text/javascript">
	function ampliarfotoseccion(nombre){
		ventana("archivos.asp?ac=ampliarfotoseccion&archivo="+nombre,'AmpliarFoto',100,100,0)
	}
	function ampliariconoseccion(nombre){
		ventana("archivos.asp?ac=ampliariconoseccion&archivo="+nombre,'AmpliarIcono',100,100,0)
	}
	</script>
<table width="100%"  border="0" cellpadding="2" cellspacing="2">
  <tr class="fondoAdmin">
    <td width="13" align="center" bgcolor="#F5F5F5"><nobr>&nbsp;&nbsp;&nbsp;</nobr></td>
    <td width="100%"><b>Nombre</b></td>
    <%if config_foto_seccion then%>
    <td align="center"><b>Foto</b></td>
<%end if%>
    <%if config_icono_seccion then%>
    <td align="center"><b>Icono</b></td>
	<%end if%>
    <%if config_activo_seccion2 then%>
    <td align="center"><nobr><b> <a href="main.asp?ac=adminsecciones2&seccion=<%=RE("S_ID")%>" class="aAdmin"><%=config_nom_secciones%></a></b></nobr></td>
    <td align="center" bgcolor="#f7f7f7">&nbsp;</td>
    <td align="center" bgcolor="#f7f7f7">&nbsp;</td>
    <%end if%>
  </tr>
  <%while not re.eof%>
  <tr>
    <td align="center" class="fondoOscuroAdmin" title="ID: <%=re("S_ID")%> ALIAS: <%=re("S_ALIAS")%> "><nobr><b><font color="#FFFFFF"><%=re("S_ORDEN")%></font></b></nobr></td>
    <td bgcolor="#FFFFFF"><%=re("S_NOMBRE")%><span title=" Número de registros en esta sección "> (<%=re("S_REGISTROS")%>)</td>
    <%if config_foto_seccion then%>
    <td align="center" bgcolor="#FFFFFF">
	<%if re("S_FOTO") <> "" then%>
	<a href="javascript:ampliarfotoseccion('<%=re("S_FOTO")%>')"><img src="img/imagen.gif" width="18" height="18" border="0"></a>
	<%end if%>
	</td>
	<%end if%>
    <%if config_icono_seccion then%>
    <td align="center" bgcolor="#FFFFFF">
	<%if re("S_ICONO") <> "" then%>
	<a href="javascript:ampliariconoseccion('<%=re("S_ICONO")%>')"><img src="img/imagen.gif" width="18" height="18" border="0"></a>
	<%end if%>
	</td>
	<%end if%>
    <%if config_activo_seccion2 then%>
    <td align="center" bgcolor="#FFFFFF"><NOBR> (<%=re("S_SUBSECCIONES")%>)</NOBR></td>
    <td align="center" <%
	if re("S_ACTIVO") then
		Response.Write "bgcolor='#95E37D' title=' Sección activada '"
	else
		Response.Write "bgcolor='#EA5940' title=' Sección desactivada '"
	end if
	%>>&nbsp;</td>
    <%end if%>
    <td align="right"><table border="0" cellspacing="0" cellpadding="0">
      <tr>
	  <%if NOT re("S_BLOQUEADA") then%>
	  	<td><a href="javascript:moverSeccion(<%=re("S_ID")%>,'subir')"><img src="img/flecha_arriba_h.gif" alt=" Subir " width="15" height="18" border="0"></a><a href="javascript:moverSeccion(<%=re("S_ID")%>,'bajar')"><img src="img/flecha_abajo_h.gif" alt=" Bajar " width="15" height="18" border="0"></a></td>
		<%end if%>
	  <%if re("S_RENOMBRAR") then%>
	  	<td><a href="javascript:editarSeccion(<%=re("S_ID")%>)"><img src="img/lapiz.gif" alt=" Editar " width="18" height="18" border="0"></a></td>
	  <%end if%>
	  <%if re("S_ELIMINAR") and re("S_REGISTROS")=0 then
	  	if config_activo_seccion then
			if config_activo_seccion2 then
				if re("S_SUBSECCIONES")=0 then%>
					<td><a href="javascript:eliminarSeccion(<%=re("S_ID")%>)"><img src="img/papelera.gif" alt=" Eliminar " width="18" height="18" border="0"></a></td>
				<%end if
			else%>
				<td><a href="javascript:eliminarSeccion(<%=re("S_ID")%>)"><img src="img/papelera.gif" alt=" Eliminar " width="18" height="18" border="0"></a></td>
			<%end if
	  end if
	  end if%>
      </tr>
    </table>      </td>
  </tr>
  <%re.movenext : wend%>
</table>
<%end if ' re.eof%>
<br>

<%	consultaXClose()
end if%>
<script>f.ac.value = "adminsecciones"</script>