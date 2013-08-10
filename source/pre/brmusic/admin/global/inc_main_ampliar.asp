<!--#include virtual="/admin/usuarios/inc_gestion_grupos.asp" -->
<%

	id = 0 + ("0" & request("id"))
	sql = "SELECT *"
	sql = sql & " FROM REGISTROS, SECCIONES"
	if config_activo_seccion2 then
		sql = sql & ", SECCIONES2"
	end if
	sql = sql & " WHERE (R_SECCION = S_ID)"
	if config_activo_seccion2 then
		sql = sql & " AND (R_SECCION2 = S2_ID)"
	end if
	sql = sql & " AND (R_ID = "& id &")"

	consultaXOpen sql,1
	if re.eof then
		unerror = true : msgerror = "No se ha encontrado el registro solicitado."
	else
		
		if cualid = "usuarios" then
			set miGrupo = setGrupo(re("R_SECCION"))
		end if
		
	end if
	if unerror then
		Response.Write("<b>Error</b><br>"& msgerror)
	else%>

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
        <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF"><nobr>Informaci&oacute;n
              extendida</nobr></font></b></td>
        <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
      </tr>
    </table>
	<br>

<%if config_creador then%>
	<table width="100%"  border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td>
	<img src="img/usuario.gif" width="18" height="18" align="absmiddle">
	&nbsp;<font color="#849ACE" size="1"><b><%=getNombreUsuario(re("R_USUARIO"))%></b>
	&nbsp;<%=re("R_AUTOFECHA")%>
	&nbsp;<%=re("R_AUTOHORA")%></font>
	<%if re("R_ULTIMO_USUARIO") <> "" then%>
	</td>
	<td align="right">
		<font color="#849ACE" size="1">&Uacute;ltima edici&oacute;n: <%="<b>"&getNombreUsuario(re("R_ULTIMO_USUARIO"))& "</b> " & re("R_ULTIMA_EDICION")%></font>
	<%end if%>
	</td>
    </tr>
</table>
<table width="100%" height="4"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="img/spacer.gif" height="4"></td></td>
  </tr>
</table>
<%end if%>
	
	<table width="100%"  border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td align="left" bgcolor="#FFFFFF"><b>Secci&oacute;n</b>: <%=re("S_NOMBRE")%><%
	if config_activo_seccion2 then
		if re("R_SECCION2") > 1 then
			Response.Write " <b>></b> " & re("S2_NOMBRE")
'			if re("R_SECCION3") > 1 then
'				Response.Write " <b>></b> " & re("S3_NOMBRE")
'			end if
		end if
	end if%></td>
	<%if config_portada and re("R_PORTADA") then%>
  	<td align="center" valign="middle" bgcolor="#FFFFFF"><span title=" Este registro aparece en portada.  "><img src="img/bandera.gif" width="18" height="18" align="absmiddle"><span class="Estilo3">EN PORTADA</span></span> </td>
	<%end if%>
	<%if config_activo then%>
		<td align="center" valign="middle" bgcolor="#FFFFFF">
		<%if re("R_ACTIVO") then%>
		  	<span title=" Este registro está activo. "><span class="Estilo3 Estilo8">ACTIVO</span></span>
		<%else%>
		  	<span title=" Este registro está inactivo. "><span class="Estilo3 Estilo9">INACTIVO</span></span>
	  <%end if%>	  </td>
	<%end if%>
  	<%if config_fecha and re("R_FECHA") <> "" and re("R_FECHA") > 0 then%>
    <td align="right" bgcolor="#FFFFFF"><%=re("R_FECHA")%></td>
	<%end if%>
    </tr>
</table>

<table width="100%" height="4"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="img/spacer.gif" height="4"></td></td>
  </tr>
</table>

	<table width="100%"  border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td><b><%=config_nom_titulo%></b></td>
    </tr>
  <tr>
    <td bgcolor="#FFFFFF"><%
	if config_idioma_bd <> "" then
		Response.Write re("R_TITULO_"& session("idioma"))
	else
		Response.Write re("R_TITULO")
	end if
	%></td>
    </tr>
</table>

<%if cualid = "usuarios" then%>
	<table width="100%"  border="0" cellpadding="4" cellspacing="0">
		<tr>
		<td><b>E-mail</b></td>
		</tr>
		<tr>
		<td bgcolor="#FFFFFF"><%=re("R_EMAIL")%></td>
		</tr>
	</table>
<%end if%>


<table width="100%" height="4"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="img/spacer.gif" height="4"></td></td>
  </tr>
</table>

<!-- Campo configurables -->

<%

	' Utilizo el nodo miGrupo que contiene seteado el grupo de la seccion escogida
	if cualid = "usuarios" then
		set nodosCampo = miGrupo.selectNodes("//grupos/grupo[@id="& re("R_SECCION") &"]//dato")
	else
		set nodosCampo = nodoCualid.childNodes
	end if


for each a in nodosCampo

	campo = ""& a.getAttribute("campo")

	if a.nodeName = "dato" and campo <> "" then
		nombrecorto = a.getAttribute("nombrecorto")
		c_titulo = a.getAttribute("titulo")
		c_tipo = a.getAttribute("tipo")

		' Campos para Idiomas
		if config_idioma_bd <> "" then
			r_campo = re("R_"& campo &"_"& session("idioma"))
		else
			r_campo = re("R_"& campo)
		end if
		
		
		select case c_tipo
		case "color"%>
		<table width="100%"  border="0" cellpadding="4" cellspacing="0">
		  <tr>
			<td><b><%=c_titulo%></b></td>
		  </tr>
		  <tr>
			<td bgcolor="#FFFFFF">
			<%if r_campo <> "" then%>
			<font face="Courier New, Courier, mono"><%=r_campo%></font><br>
			<table width="40" height="40" bgcolor="<%=r_campo%>" border="0" cellspacing="0" cellpadding="0">
			<tr>
			<td>&nbsp;</td>
			</tr>
			</table>

			<%else%>
			&nbsp;
			<%end if%></td>
		  </tr>
		</table>
		<%case else%>
		<table width="100%"  border="0" cellpadding="4" cellspacing="0">
		  <tr>
			<td><b><%=c_titulo%></b></td>
		  </tr>
		  <tr>
			<td bgcolor="#FFFFFF"><%if r_campo <> "" then Response.Write replace(r_campo,vbCrlf,"<br>") else Response.Write "&nbsp;" end if%></td>
		  </tr>
		</table>
		<%end select%>
		
		<table width="100%" height="4"  border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td><img src="img/spacer.gif" height="4"></td></td>
		  </tr>
		</table>
	<%end if
next%>

<%if config_foto then%>
	<table width="100%"  border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td colspan="2"><b><%=config_nom_foto%></b></td>
    <td width="150" valign="bottom">
	<%if config_posicion_foto then%>
	<b>Posici&oacute;n</b>
	<%end if%>
	</td>
  </tr>
	<%if ""&re("R_FOTO")<>"" then%>
  <tr>
    <td width="20" align="center" valign="middle" bgcolor="#FFFFFF"><a href="javascript:ampliarfoto('<%=re("R_FOTO")%>')"><img src="img/imagen.gif" alt=" Ver foto ampliada. " width="18" height="18" border="0" align="absmiddle"></a></td>
    <td bgcolor="#FFFFFF"><%=re("R_FOTO")%></td>
    <td bgcolor="#FFFFFF">
	<%if config_posicion_foto then%>
	<%=re("R_POS_FOTO")%>
	<%end if%></td>
  </tr>
	<%else%>
  <tr>
    <td colspan="2" valign="middle" bgcolor="#FFFFFF">Vacio</td>
    <td valign="middle" bgcolor="#FFFFFF">
	<%if config_posicion_foto then%>
	<%=re("R_POS_FOTO")%>
	<%end if%></td>
  </tr>
	<%end if%>
	<%if config_pie_foto then%>
  <tr>
    <td colspan="3" valign="middle" bgcolor=""><b>Pie de foto</b></td>
    </tr>
  <tr>
    <td colspan="3" valign="middle" bgcolor="#FFFFFF"><%=re("R_PIE_FOTO")%></td>
    </tr>
	<%end if%>
</table>
<%end if%>

<%if config_icono then%>
	<table width="100%"  border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td><b><%=config_nom_icono%></b></td>
    <td width="150">
	<%if config_posicion_icono then%>
	<b>Posici&oacute;n</b>
	<%end if%></td>
  </tr>
	<%if ""&re("R_ICONO")<>"" then%>
  <tr>
    <td valign="middle" bgcolor="#FFFFFF"><img src="../../datos/<%=session("idioma")%>/<%=session("cualid")%>/iconos/<%=re("R_ICONO")%>" border="0"><br>
      <%=re("R_ICONO")%></td>
    <td valign="bottom" bgcolor="#FFFFFF">
	<%if config_posicion_icono then%>
	<%=re("R_POS_ICONO")%>
	<%end if%></td>
  </tr>
	<%else%>
  <tr>
    <td valign="middle" bgcolor="#FFFFFF">Vacio</td>
    <td valign="middle" bgcolor="#FFFFFF">
	<%if config_posicion_icono then%>
	<%=re("R_POS_ICONO")%>
	<%end if%></td>
  </tr>
	<%end if%>
</table>
<%end if%>

<%if config_fuente then%>
	<table width="100%"  border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td><b><%=config_nom_fuente%></b></td>
    </tr>
  <tr>
    <td bgcolor="#FFFFFF">
	<%
	if ""&re("R_FUENTE")<> "" then
		Response.Write re("R_FUENTE")
	else
		Response.Write "&nbsp;"
	end if
	
	%></td>
    </tr>
</table>
<br>

	<table width="100%"  border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td><b><%=config_nom_enlace%></b></td>
    </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<%
	if re("R_ENLACE") <> "" then
		Response.Write re("R_ENLACE")
	else
		Response.Write "&nbsp;"
	end if
%></td>
    </tr>
</table>
<br>


<%end if

if config_editar then%>
<table width="100%" height="4"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="img/spacer.gif" height="4"></td></td>
  </tr>
</table>

<table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td align="left"><font color="#666666" size="1">ID: <%=re("R_ID")%></font></td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="left">&nbsp;</td>
    <td align="right"><table border="0" cellpadding="2" cellspacing="0" onClick="editar(<%=id%>)" class="botonAdmin" title=" Editar este registro ">
      <tr>
        <td>Editar</td>
      </tr>
    </table></td>
  </tr>
</table>
<%end if

	end if
	consultaXClose()
%>