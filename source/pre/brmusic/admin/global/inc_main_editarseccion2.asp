<%	nombre = request.Form("nombre")
	largo = len(nombre)
	foto = ""&request.Form("foto")
	icono = ""&request.Form("icono")
	if nombre <>"" then
		if largo > config_maxcarseccion then
		%><script language="javascript" type="text/javascript">
			alert(" * Error en longitud de nombre *\n\nEl nombre de sección escrito tiene <%=largo%> caracteres.\nPor favor escriba un nombre igual o inferior a <%=config_maxcarseccion%> caracteres.\n")
			location.href="main.asp?ac=adminsecciones"
		</script><%
		else
			id = request.Form("id")
			nombre = ""& request.Form("nombre")

			set secciones = nodoCualid.selectSingleNode("secciones2")
			if typeOK(secciones) then
				camposconfigurables = true
			else
				camposconfigurables = false
			end if

			if not unerror then
				sql = "UPDATE SECCIONES2 SET "
		
				' Campos fijos
				sql = sql & "S2_NOMBRE = '"& replace(nombre,"'","''") &"'"
				if config_activo_seccion then
					sql = sql & ", S2_ACTIVO = "& numero(request.Form("activo"))
				end if
	
				' Campos configurables
				if camposconfigurables then
					for each a in secciones.childNodes
						if a.nodeName = "campo" then
							sql = sql & ", S2_"& ucase(a.getAttribute("nombre")) &"= '"& replace(request.Form(a.getAttribute("nombre")),"'","''") &"'"
						end if
					next
				end if
				sql = sql & " WHERE S2_ID = " & id
				
				if exeSql(sql,conn_) <> "" then
					unerror = true : msgerror = "Se ha producicdo un error en la ejecución SQL.<br>SQL: " & sql
				end if

			end if ' unerror

			if not unerror then
				redir = ""
				select case foto
					case "anadir", "cambiar", "auto"
						redir = "archivos_frames.asp?ac=formguardarfotoseccion2&id="& id &"&icono="& icono
					case "eliminar"
						redir = "archivos_frames.asp?ac=quitarfotoseccion2&id="& id &"&icono="& icono
					case "nada"
						select case icono
							case "anadir","cambiar"
								redir = "archivos_frames.asp?ac=formguardariconoseccion2&id="& id
							case "eliminar"
								redir = "archivos_frames.asp?ac=quitariconoseccion2&id="& id
						end select
				end select
				if redir <> "" then
					Response.Redirect(redir)
				else
					Response.Redirect("main.asp?ac=adminsecciones2&msg=Sección editada correctamente.")
				end if
			end if
			
			if unerror then
				Response.Write "<b>Error</b><br>" & msgerror
			end if

		end if' largo


	else
		id = request.QueryString("id")
		if id <> "" and isNumeric(id) then
	
		sql = "SELECT * FROM SECCIONES2 WHERE S2_ID = " & id
		consultaXOpen sql,1
%>
<script language="javascript" type="text/javascript">
	function ampliarfotoseccion2(nombre){
		ventana("archivos.asp?ac=ampliarfotoseccion2&archivo="+nombre,'AmpliarFoto',100,100,0)
	}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Editar <%=config_nom_secciones2%></font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>Escriba el nuevo nombre para la secci&oacute;n <b><%=re("S2_NOMBRE")%></b> y pulse en el bot&oacute;n aceptar.<br>
	Recuerde que el nombre debe ser de <%=config_maxcarseccion%> car&aacute;teres como m&aacute;ximo.</td>
    
<%if config_activo_seccion then%>
	<td align="right" valign="top"><table  border="0" cellspacing="0" cellpadding="2">
      <tr>
        <td><input type="checkbox" name="activo" id="activo" value="1" <%if re("S2_ACTIVO") then Response.Write "checked" end if%>></td>
        <td><label for="activo"><b>Activo</b></label></td>
      </tr>
    </table></td>
<%end if%>

  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="2" cellspacing="0">
  <tr>
    <td align="left"><b>T&iacute;tulo*:</b></td>
    </tr>
  <tr>
    <td align="left"><input name="nombre" type="text" class="campoAdmin" id="nombre" style="width:100%" value="<%=re("S2_NOMBRE")%>" maxlength="100"></td>
    </tr>
</table>
<table cellpadding="2" cellspacing="0" border="0" width="100%">
<%
numCampo = 0
set nodoSecciones2 = nodoCualid.selectSingleNode("secciones2")
if not typeOK(nodoSecciones2) then
	unerror = true : msgerror = "No se ha encontrado el nodo se configuración para las secciones."
end if

if not unerror then
	for each a in nodoSecciones2.childNodes
		c_titulo = a.getAttribute("titulo")
		c_nombre = a.getAttribute("nombre")
		c_tipo = a.getAttribute("tipo")
		numCampo = numCampo + 1
		if a.nodeName = "campo" then
	%>
	  <tr>
		<td colspan="2"><b><%=c_titulo%></b><%if ""&a.getAttribute("requerido")="1" then%>*<%end if%></td>
	  </tr>
	  <tr>
		<td colspan="2" valign="top">
		  <%select case c_tipo
		  case "text"%>
		  <input name="<%=c_nombre%>" type="text" <%if a.getAttribute("manipulable")="0" then response.Write("disabled='true'") end if%> class="campoAdmin" id="subtitulo" style="width:100%" value="<%=re("S2_" & c_nombre)%>" maxlength="255">
		  <%case "memo"%>
		  <textarea name="<%=c_nombre%>" rows="<%=a.getAttribute("filas")%>" wrap="virtual" class="areaAdmin" style="width:100%"><%=re("S2_" & c_nombre)%></textarea>
		<%if a.getAttribute("editorhtml") = 1 then%>
		  <script language="javascript1.2">
		  editor_generate("<%=c_nombre%>");
		  </script>
		  <%end if%>
		  <%case "combo"%>
		  <select name="<%=c_nombre%>" class="campoAdmin">
		  <option value="">Seleccione una ...</option>
		  <%for each opcion in a.childNodes%>
			<option value="<%=opcion.getAttribute("valor")%>" <%if re("S2_" & c_nombre) = opcion.getAttribute("valor") then Response.Write "selected" end if%>><%=opcion.getAttribute("titulo")%></option>
		  <%next%>
		  </select>
	
		  <%case "opcion"
		  	inc = 0
			for each opcion in a.childNodes
			inc = inc + 1%>
			<input type="radio" name="<%=c_nombre%>" id="<%=c_nombre&inc%>" value="<%=opcion.getAttribute("valor")%>" <%if re("S2_" & c_nombre) = opcion.getAttribute("valor") then Response.Write "checked" end if%>><label for="<%=c_nombre&inc%>"><%=opcion.getAttribute("titulo")%></label>
			<%next

		case "check"
			valores = split(""&re("S2_" & c_nombre),",")
			for each valor in valores
				cadena = cadena &"|"& trim(valor)
			next
			cadena = cadena &"|"
			inc = 0
			for each opcion in a.childNodes
			inc = inc + 1%>
			<input  type="checkbox" name="<%=c_nombre%>" id="<%=c_nombre&inc%>" value="<%=opcion.getAttribute("valor")%>" <%if inStr(cadena,opcion.getAttribute("valor"))>0 then Response.Write "checked" end if%>><label for="<%=c_nombre&inc%>"><%=opcion.getAttribute("titulo")%></label>
			<%next
			
		case "color"%>
			<div align="center">
			<input name="<%=c_nombre%>" type="text" class="campoAdmin" value="<%=re("S2_" & c_nombre)%>" size="8" maxlength="7" readonly="true">
			<div id="capaColorFlash" style="position:absolute; z-index:<%=10-numCampo%>;width: 172; height: 155; visibility: visible;">
			<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="172" height="155">
			<param name="movie" value="color.swf?nombre=<%=c_nombre%>&col=<%=Replace(re("S2_" & c_nombre),"#","")%>">
			<param name="quality" value="high">
			<param name="WMODE" value="transparent">
			<embed src="color.swf?nombre=<%=c_nombre%>&col=<%=Replace(re("R_" & c_nombre),"#","")%>" width="172" height="155" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent"></embed>
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
	
	case else%>
		  <font color="#FF0000">Error en definici&oacute;n del XML.</font>
		  <%end select%>
		</td>
	  </tr>
	  <%end if
	  next
end if  
  %>
  </table>



<br>

<%if config_foto_seccion2 then%>
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
      <tr>
        <td><b>Foto</b></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF">
	<%if re("S2_FOTO") <> "" then%>
			<a href="javascript:ampliarfotoseccion2('<%=re("S2_FOTO")%>')"><img src="img/imagen.gif" width="18" height="18" border="0"></a>
			<input name="foto" id="foto_cambiar" type="radio" value="cambiar"><label for="foto_cambiar">Cambiar</label>
			<input name="foto" id="foto_eliminar" type="radio" value="eliminar"><label for="foto_eliminar">Eliminar</label>
			<input name="foto" id="foto_nada" type="radio" value="nada" checked><label for="foto_nada">Nada</label>
		<%else%>
			<input name="foto" id="foto_anadir" type="radio" value="anadir"><label for="foto_anadir">A&ntilde;adir</label>
			<input name="foto" id="foto_nada" type="radio" value="nada" checked><label for="foto_nada">No a&ntilde;adir</label>
		<%end if%>
	</td>
      </tr>
    </table>
<%end if%>
<%if config_icono_seccion2 then%>
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
      <tr>
        <td><b>Icono</b></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF">
		<%if re("S2_ICONO") <> "" then%>
		<img src="../../datos/<%=session("idioma")%>/<%=cualid%>/iconos_seccion2/<%=re("S2_ICONO")%>">
		<br>
		<input name="icono" id="icono_anadir" type="radio" value="auto">
        <label for="icono_anadir">Auto</label>
        <input name="icono" id="icono_cambiar" type="radio" value="cambiar"><label for="icono_cambiar">Cambiar</label>
		<input name="icono" id="icono_eliminar" type="radio" value="eliminar"><label for="icono_eliminar">Eliminar</label>
		<input name="icono" id="icono_nada" type="radio" value="nada" checked><label for="icono_nada">Nada</label>
	<%else%>
		<input name="icono" id="icono_auto" type="radio" value="auto"><label for="icono_auto">Auto</label>
		<input name="icono" id="icono_anadir" type="radio" value="anadir"><label for="icono_anadir">A&ntilde;adir</label>
		<input name="icono" id="icono_nada" type="radio" value="nada" checked><label for="icono_nada">No a&ntilde;adir</label>
	    <%end if%>
</td>
      </tr>
    </table>
<br>
<%end if%>
	<script>
f.id.value=<%=id%>
</script>

<table width="100%" border="0" cellpadding="2" cellspacing="0">
  <tr>
    <td align="right"><table border="0" cellspacing="0" cellpadding="1">
      <tr>
        <td><input name="" type="submit" class="botonAdmin" onClick="location.href='main.asp?ac=adminsecciones'" value="Volver">
        </td>
        <td><input name="aceptar_btn" type="submit" disabled="true" class="botonAdmin" id="aceptar_btn" value="Aceptar">
        </td>
      </tr>
    </table>      </td>
    </tr>
</table>
<script language="javascript" type="text/javascript">
<!--
function cargada(){
	f.aceptar_btn.disabled = false
}
window.onload = cargada;
//-->
</script>
<br>
<br>
<%
			consultaXClose()
		else
			Response.Redirect("inicio.asp")
		end if
	end if%>