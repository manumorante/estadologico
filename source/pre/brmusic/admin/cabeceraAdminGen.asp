<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<%usuariologeado = true%>
<!--#include file="usuarios/rutinasParaAdmin.asp" -->
<html>
<head>
<title>Cabecera de administraci&oacute;n</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="global/estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
a:link {
	color: #FFFFFF;
	text-decoration: none;
}
a:visited {
	color: #FFFFFF;
	text-decoration: none;
}
a:hover {
	color: #FFFFFF;
	text-decoration: none;
}
a:active {
	color: #FFFFFF;
	text-decoration: none;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>
<body>
<%
if not unerror then
	Dim ruta_xml_config ' RUTA ^
	ruta_xml_config = "/datos/xml_admin_config.xml"

	' *** CARGA DEL XML DE CONFIGURACIÓN *** (xml de configuración general para todas la cualidades del sistema en la página actual)
	set xml_config = CreateObject("MSXML.DOMDocument")
	if not xml_config.Load(Server.MapPath(ruta_xml_config)) then
		unerror = true : msgerror = "'XML config' Error de carga."
	else
		set nodoConfig = xml_config.selectSingleNode("configuracion")
		if typeName(nodoConfig) = "Nothing" or typeName(nodoConfig) = "Empty" then
			unerror = true : msgerror = "'XML config' Error de estructura."
		end if
	end if
end if

if not unerror then
	' VARIABLE DEL SITIO
	Dim carpetasitio, nombresitio
	carpetasitio = ""& nodoConfig.getAttribute("carpetasitio")
	nombresitio = ""& nodoConfig.getAttribute("nombresitio")
	if carpetasitio = "" then
		'unerror = true : msgerror = "No se ha declarado la carpeta de este sitio web."
	end if
	
	if nombresitio = "" then
		'unerror = true : msgerror = "No se ha declarado el nombre de este sitio web."
	end if
end if
	
if unerror then
	if coderror = 11 then
		%>
		<script>
			// alert("Usuario no logeado.")
			top.location.href='usuarios/validar.asp?msg=<%=msgerror%>'
		</script>
		<%
		Response.Write "<b>Error:</b><br>"& msgerror
	else
		Response.Write "<b>Error:</b><br>"& msgerror
	end if
else%>
	<script>

		// Ir 
		function goTo(cualid) {
		
			if (cualid != ""){
				if (cualid == "edicion") {
					var ir = "edicion/index.asp?cualid=edicion"
				} else if (cualid == "edicionmovil") {
					var ir = "edicion/index.asp?cualid=edicionmovil"
				} else if (cualid == "principal") {
					var ir = "inicio.asp"
				} else if (cualid == "carrito") {
					var ir = "carrito/index.asp"
				} else if (cualid.indexOf("admindatos_")>=0) {
					var ir = "admindatos/index.asp?cualid="+cualid
				} else {
					var ir = "global/default.asp?cualid="+cualid
				}
				try{
					parent.frames[1].location.href = ir
				}catch(unerror){
					alert("No se ha encontrado el frame.")
					top.location.href='default.asp'
				}
			}else{
				//alert("El XML no esta bien definido.")
			}
		}
	function reparar(cualid) {
		
			if (cualid != ""){
				if (cualid == "edicion") {
					alert("Esta cualidad no puede ser reparada")
				} else {
					var ir = "gestionmdb/actualizardb.asp?cualid="+cualid
					try {
						parent.frames[1].location.href = ir
					}catch(unerror){
						alert("No se ha encontrado el frame.")
						top.location.href='default.asp'
					}
				}
			}else{
				// alert("El XML no esta bien definido.")
			}
		}
		//
		function salirOver() {
			imgsalir.src = "images/desconectar_on.gif"
		}
	
		//
		function salirOut() {
			imgsalir.src = "images/desconectar_off.gif"
		}
	
		//
		function desconectar() {
			if (confirm("¿Desea desconectarse del sistema?")) {
				top.location.href="usuarios/validar.asp?ac=desconectar"
			}
		}
		function principal(){
			top.frames[1].location='inicio.asp'
		}
		
		<%if session("cualid")<>"" and request("direct")="" then%>
		goTo("<%=session("cualid")%>")
		<%end if%>
	</script>
	
	<table width="100%" border="0" cellpadding="0" cellspacing="0" background="images/fondo.gif">
	<tr>
	<td width="140"><table border="0" cellpadding="3" cellspacing="0">
      <tr valign="top">
        <td width="300" align="left"><nobr><span class="ColorBlanco"><strong>Administrar:</strong></span>
              <select name="combo" onChange="goTo(this.value);" class="campoAdmin">
                <option value="principal">Principal</option>
                <%for each a in nodoConfig.childNodes
		if inStr(a.nodeName,"acceso_")<=0 then
			if getPermiso(a.nodeName, session("idioma")) then%>
                <option value="<%=a.nodeName%>" <%if session("cualid") = a.nodeName then%>selected<%end if%>><%=a.getAttribute("nombre")%></option>
                <%end if
		end if
	next%>
              </select>
            <a href="#" onClick="goTo(combo.value);return false;"><img src="images/ir.gif" alt=" Ir " width="18" height="18" border="0" align="absmiddle"></a></nobr></td>
        <% if session("usuario")=1 then %>
        <% end if %>
      </tr>
    </table></td>
	<td align="left" background="images/fondo.gif">
	<%
	
	' Pintar la liesta de idiomas disponbles
	set idiomas = setIdiomas()
	for each a in idiomas.childNodes
		if a.nodeName = session("idioma") then%>
			<span class="colorBlanco" title=" Idioma acual. "><b><%=a.text%></b></span>&nbsp;
		<%else%>
			<a href="usuarios/validar.asp?ac=cambiaridioma&nuevoidioma=<%=a.nodeName%>" class="ColorBlanco"><%=a.text%></a>&nbsp;
		<%end if
	next%>
	
<!--	<font color="#FFFFFF"><%'=getDatoUsu(1,"nombre")%></font> -->
	<td align="right">&nbsp;</td>
	<td width="30" align="center" valign="middle">&nbsp;</td>
	<td width="32" align="center" valign="middle"><a href="#" onClick="desconectar();return false;"><img src="images/desconectar_off.gif" alt=" DESCONECTAR " name="imgsalir" width="32" height="32" border="0" id="imgsalir" onMouseOver="salirOver()" onMouseOut="salirOut()"></a></td>
	</tr>
</table>

<%end if%>
</body>
</html>
