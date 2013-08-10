<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<!--#include file="inc_rutinas.asp" -->
<%
	Dim unerror
	Dim msgerror
	Dim id

	id = numero(request.QueryString("id"))
	if id <= 0 then
		unerror = true : msgerror = "No se ha recibido un id."
	end if

	if not unerror then
		ruta_xml_config = Server.MapPath("/"& c_s &"datos/xml_admin_config.xml")
		set xml_config = CreateObject("MSXML.DOMDocument")
		if not xml_config.Load(ruta_xml_config) then
			unerror = true : msgerror = "Hay un problema con el archivo de configuración general.<br>Conpruebe su ubicación y que no contenga ningún error.<br>Ruta:"& ruta_xml_config
		else
			set cualidades = xml_config.selectSingleNode("configuracion")
			if not typeOK(cualidades) then
				unerror = true : msgerror = "Hay un problema con el archivo de configuración general.<br>Conpruebe que no contenga ningún error."
			end if
		end if
	end if
	
	' Cargar la MDB
	if not unerror then%>
		<!--#include file="inc_conn.asp" -->
	<%end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Usuarios personalizados</title>
<link href="../estilos.css" rel="stylesheet" type="text/css">
<link href="estilos.css" rel="stylesheet" type="text/css">
</head>

<body class="bodyAdmin">

<%if not unerror then%>

<%select case request("ac")%>

<%case "editar"

if ""& request.QueryString("editar") = "1" then
	
	' Crear el XML de permisos
	dim str_xml
	d = chr(34)
	str_xml = "<?xml version="&d&"1.0"&d&" encoding="&d&"iso-8859-1"&d&"?>"
	str_xml = str_xml & "<permisos>"
	for each campo in request.Form()
		if campo <> "" then
			str_xml = str_xml & "<"&campo&"/>"
		end if
	next
	str_xml = str_xml & "</permisos>"

	conn_activa.execute "UPDATE REGISTROS SET R_PERMISOS = '"& str_xml &"' WHERE R_ID = "& id
'	Response.Write "str_xml: " & replace(server.HTMLEncode(str_xml),vbCrlf,"<br>")
	%>

	<script language="javascript" type="text/javascript">
	<!--
		window.close()
	//-->
	</script>
	<%else


	dim re
	dim reTotal	
	reTotal = 0
	sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
	consultaXopen sql,1
	
	if reTotal <= 0 then
		unerror = true : msgerror = "No se ha encontrado el usuario."
	end if
	
	' Cargo el XML de permiso desde el campo R_PERMISOS en la tabla
	if not unerror then
		hay_nodo_permisos = false
		if ""&re("R_PERMISOS") <> "" then
			hay_nodo_permisos = true
			set nodo_permisos = CreateObject("MSXML.DOMDocument")
			if nodo_permisos.LoadXML(re("R_PERMISOS")) then
				set nodo_permisos = nodo_permisos.selectSingleNode("permisos")
				if not typeOK(nodo_permisos) then
					unerror = true : msgerror = "El XML obtenido de la base de datos no es correcto. Le falta el nodo permisos."
				end if
			else
				unerror = true : msgerror = "No se ha podido cargar el XML contenido en la MDB."
			end if
		end if
	end if

	if not unerror then
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><nobr><b><font color="#FFFFFF">Permisos para usuarios personalizados</font></b></nobr></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>


	<table width="100%"  border="0" cellspacing="0" cellpadding="1">
		<tr>
		  <td>&nbsp;</td>
	  </tr>
		<tr>
		<td>Escoja las cualidades que desea para este usuario.<br>
		Los campos marcados con uns asterisco (*) son obligatorios.<br>
		<br>
		<form name="f" action="usuarios_personalizados.asp?ac=editar&id=<%=id%>&editar=1" method="post">
		<table width="100%"  border="0" cellpadding="1" cellspacing="0">
		  <tr bordercolor="#849ACE" class="fondoOscuroAdmin">
		    <td height="30" colspan="2"><span class="colorBlanco">&nbsp;&nbsp;Propiedades del usuario </span></td>
	      </tr>
		  <tr>
		    <td colspan="2">&nbsp;</td>
	      </tr>
		  <tr class="campoAdmin">
            <td>&nbsp;Cualidad&nbsp;</td>
            <td>&nbsp;Idiomas/Grupos</td>
          </tr>

		<%for each cualid in cualidades.childNodes
		
			if cualid.nodeName <> "idiomas" and cualid.nodeName <> "infogeneral" and cualid.nodeName <> "usuarios" then
				if claseFila = "clasefila1" then
					claseFila = "clasefila2"
				else
					claseFila = "clasefila1"
				end if%>
				
				<tr class="<%=claseFila%>">
				<td>&nbsp;<%=cualid.getAttribute("nombre")%></td>
				<td>
				<%
				select case cualid.nodeName
					
				' RESTO DE CUALIDADES
				case else
					for each idi in cualid.getElementsByTagName("idioma")
						if hay_nodo_permisos then
							set temp = nodo_permisos.selectSingleNode(cualid.nodeName &"_"& idi.getAttribute("nombre"))
							if typeOK(temp) then
								checked = "checked"
							else
								checked = ""
							end if
						end if%>
						<input id="<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>" name="<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>" type="checkbox" <%=checked%> value="1"><label for="<%=cualid.nodeName%>_<%=idi.getAttribute("nombre")%>"><%=getNombreIdioma(idi.getAttribute("nombre"))%></label>
					<%next
				end select%>&nbsp;</td>
				</tr>
			<%end if
		next%>
		<tr><td colspan="3" align="right">&nbsp;</td></tr>
		<tr class="fondoOscuroAdmin">
		  <td colspan="3" align="right"><input name="" type="button" class="botonAdmin" onClick="window.close();" value="Cancelar">
		    <input type="submit" class="botonAdmin" value="Editar"></td></tr>
		</table>
		</form>
		
		</td>
		</tr>
	</table>
	<%
	end if

	consultaXClose()
	end if
end select

end if

if unerror then%>
	<b>Error</b>:<br><%=msgerror%>
<%end if

set conn_activa = Nothing%>
</body>
</html>
