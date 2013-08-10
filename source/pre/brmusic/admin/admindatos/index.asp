<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/admin/inc_rutinas.asp" -->
<%zona = 1%>
<!--#include virtual="/admin/usa.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Formularios XML</title>
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
<link href="../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body class="bodyAdmin">

<%
dim cualid				' Cualidad.
dim ruta_xml_config		' Ruta del XML de configuracion general /datos/xml_admin_config.xml.
dim xml_config			' Nodo padre.
dim nodoForm			' Nodo de la cualidad solicitada.
dim idioma				' Idioma

if ""& request.QueryString("cualid") <> "" then
	cualid = request.QueryString("cualid")
	session("cualid") = cualid
elseif ""& session("cualid") <> "" then
	cualid = session("cualid")
else
	unerror = true : msgerror = "No se ha recibido una cualidad."
end if

idioma = ""&session("idioma")
if idioma = "" then
	unerror = true : msgerror = "Se ha perdido el idioma."
end if

' Cargo el XML de configuración
if not unerror then
	ruta_xml_config = "/" & c_s & "datos/xml_admin_config.xml"
	set xml_config = CreateObject("MSXML.DOMDocument")
	if not xml_config.Load(Server.MapPath(ruta_xml_config)) then
		unerror = true : msgerror = "'XML config' Error de carga."
	else
		set nodoForm = xml_config.selectNodes("//"& cualid).item(0)
		if not typeOK(nodoForm) then
			unerror = true : msgerror = "No se ha encontrado la cualidad solicitada.<br>Cualidad: " & ucase(replace(cualid,"admindatos_",""))
		end if
	end if
end if

' Cargo el XML de datos
if not unerror then
	ruta_xml_datos = "/" & c_s & "datos/"& idioma &"/"& cualid &"/"& cualid &".xml"
	set xml_datos = CreateObject("MSXML.DOMDocument")
	if not xml_datos.Load(Server.MapPath(ruta_xml_datos)) then
'		unerror = true : msgerror = "'XML datos' Error de carga."
	else
		set nodoDatos = xml_datos.selectSingleNode("datos")
'		if not typeOK(nodoDatos) then
'			unerror = true : msgerror = "No se ha encontrado el nodo Datos en el XML de datos."
'		end if
	end if
end if

if not unerror then
	titulo = ""& nodoForm.getAttribute("nombre")
	descripcion = ""& nodoForm.getAttribute("descripcion")
	
	action = "index.asp?ac=enviar"
	idioma = "esp"

	Select case request.QueryString("ac")
	case "enviar"

		dim nuevo_xml_datos		' Nuevo XML que crearemos.
		dim nodo_total_datos	' El nodo 'padre' de nuevo XML.
		dim nodo_temp			' Nodo temporal para los hijos.

		set nuevo_xml_datos = CreateObject("MSXML.DOMDocument")
		set nuevo_nodo_datos = nuevo_xml_datos.createElement("datos")
	
		set nodo_total_datos = xml_config.selectNodes("/configuracion/"& cualid &"//dato")
		for each nodo in nodo_total_datos
			nombrecorto = ""& nodo.getAttribute("nombrecorto")
			if nombrecorto <> "" then
				set nodo_temp = nuevo_xml_datos.createElement(nombrecorto)
				nuevo_nodo_datos.appendChild(nodo_temp)
				nodo_temp.text = ""& request.Form(nombrecorto)
			end if
		next
		nuevo_xml_datos.appendChild(nuevo_nodo_datos)
		set nodo_temp = Nothing
		set nodo_total_datos = Nothing
		set xml_datos = Nothing
		
		' Guardar XML
		on error resume next
		nuevo_xml_datos.save Server.MapPath("/" & c_s & "datos/"& idioma &"/"& cualid &"/"& cualid & ".xml")
		if err <> 0 then
			unerror = true : msgerror = "Se ha producido un error al guardar.<br>"& err.description
		end if
		on error goto 0
		

		if not unerror then
			Response.Redirect("index.asp?msg=Cambio realizado")
		end if
		
		if unerror then
			Response.Write "<b>Error:</b><br>"& msgerror
		end if

	case else%>

		
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
			<td width="8" height="19"><img src="../global/img/titulo_izq.gif" width="8" height="19"></td>
			<td align="center" valign="middle" background="../global/img/titulo_cen.gif"><nobr><b><font color="#FFFFFF"><%=titulo%></font></b></nobr></td>
			<td width="8" height="19"><img src="../global/img/titulo_der.gif" width="8" height="19"></td>
			</tr>
		</table>

		<table width="100%"  border="0" cellspacing="0" cellpadding="1">
			<tr>
			<td>&nbsp;</td>
			</tr>

			<%if descripcion <> "" then%>
				<tr>
				<td align="center"><%=descripcion%></td>
				</tr>
			<%end if%>

			<tr>
			  <td>&nbsp;</td>
		  </tr>
			<tr>
			  <td><div align="center">
		        <font color="#009900"><b>
		        <%
			  if ""&request.QueryString("msg") <> "" then
			  	Response.Write request.QueryString("msg")
			  end if
			  %>
			    </b></font></div></td>
		  </tr>
			<tr>
			  <td><%=leeForm (nodoForm, nodoDatos, "1", action, idioma)%></td>
		  </tr>
		</table>
			
			</td>
          </tr>
</table>

	<%end select

end if ' unerror%>

<%if unerror then%>
<b>Error: </b><%=msgerror%>
<%end if%>
</body>
</html>
