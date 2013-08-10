<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<!--#include file="inc_seguridad.asp" -->
<!--#include file="xelupload.asp" -->
<!--#include file="inc_conn.asp" -->
<!--#include virtual="/admin/inc_rutinas.asp" -->
<!--#include file="inc_rutinas.asp" -->
<!--#include file="inc_inicia_xml.asp" -->
<%
idioma = session("idioma")
cualid = session("cualid")
inicia_xml
Dim id, up, fich, rutaDatos
Dim maximo, ruta, nombre
dim reTotal, re
maximo = 5242880 ' 5 Mb
rutaDatos = "/"& c_s &"datos/"& session("idioma") &"/"& session("cualid")

icono_anchomax = 150

%>
<!--#include virtual="/admin/visores/inc_conn.asp" -->
<html>
<head>
<title>Imagen</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="estilos.css" rel="stylesheet" type="text/css">
<script>
function tamano(w,h) {
	try{
		if (w < screen.width-100 && h < screen.height-100) {
			var winl = (screen.width - w) / 2;
			var wint = (screen.height - h) / 2;
			moveTo(winl,wint)
			window.resizeTo(w+30,h+60)
		} else {
			w = screen.width - 100
			h = screen.height - 150
			var winl = (screen.width - w) / 2;
			var wint = (screen.height - h) / 2;
			moveTo(winl,wint)
			window.resizeTo(w+30,h+60)
			alert("La imagen es mas grande que la pantalla y se ve cortada.")
		}
	} catch(unerror){
		alert(unerror.description + "\nPara ver la imagen completa cierre y vuelva a ampliar.")
	}
}

</script>
</head>
<body class="bodyAdmin">
<%select case request.QueryString("ac")

case "generar_icono"
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Crear icono desde foto</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<script>
function numerico(c){
	if (isNaN(c.value)){
		alert("Escriba un valor numérico.")
		c.focus()
	}
}
</script>
<form name="f" method="post" action="">
<input type="hidden" name="anchomax" value="<%=anchomax_cuerpo_sitio%>">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><table  border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
      <tr>
        <td align="center"><img src="<%=request.QueryString("archivo")%>" name="id_img" id="id_img"></td>
      </tr>
    </table>

      <table width="80%"  border="0" align="center" cellpadding="2" cellspacing="2">
        
        <tr align="left" valign="top" bgcolor="#FFFFFF">
          <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
                <td><font color="#0066CC">Tama&ntilde;o foto (p&iacute;xeles):</font></td>
                <td align="right"><b>Ancho</b>:
                  <input name="ancho_foto" type="text" class="campoAdmin" id="ancho_foto" size="5" value="">
                  <a href="JavaScript:refrescar('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>  <b>Alto</b>:
                  <input name="alto_foto" type="text" class="campoAdmin" id="alto_foto" size="5" value="">
                  <a href="JavaScript:refrescar('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a></td>
                <td align="right">&nbsp;</td>
              </tr>
          </table></td>
        </tr>
        
        <tr valign="top">
		<%if ""&request.QueryString("icono") = "auto" then%>
          <td align="center" valign="middle"><table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#849ACE">
            <tr>
              <td><table width="100%"  border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                <tr>
                  <td align="center"><table  border="0" cellspacing="2" cellpadding="3">
                      <tr align="left">
                        <td colspan="2"><font color="#0066CC">Tama&ntilde;o icono (p&iacute;xeles):</font></td>
                      </tr>
                      <tr>
                        <td align="right"><b>Ancho</b>: </td>
                        <td bgcolor="#F5F5F5">&nbsp;
                            <input name="ancho_icono" type="text" class="campoAdmin" id="ancho_icono" onBlur="numerico(this)" size="5">
                            <a href="JavaScript:refrescarIcono('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>&nbsp;</td>
                      </tr>
                      <tr>
                        <td align="right"><b>Alto</b>:</td>
                        <td bgcolor="#F5F5F5">&nbsp;
                            <input name="alto_icono" type="text" class="campoAdmin" id="alto_icono" onBlur="numerico(this)" size="5">
                            <a href="JavaScript:refrescarIcono('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>&nbsp;</td>
                      </tr>
                  </table></td>
                  <td align="center"><table  border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
                      <tr>
                        <td><img src="<%=request.form("archivo")%>" name="id_img_icono" width="156" hspace="3" vspace="3" id="id_img_icono"></td>
                      </tr>
                  </table></td>
                </tr>
              </table></td>
            </tr>
          </table>
            </td>
          </tr>
		  
		  <%end if%>
		  
        
      </table>
	  </td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
    <input name="enviar" type="button" class="botonAdmin" id="enviar" onClick="refrescar('enviar');" value="Enviar"></td>
  </tr>
</table>
</form>
<script language="javascript" type="text/javascript">
	// Temp Globales
	img = document.getElementById("id_img")
	<%if ""&request.QueryString("icono") = "auto" then%>
		img_icono = document.getElementById("id_img_icono")
		var	ancho_icono_tmp
		var alto_icono_tmp
	<%end if%>
	<%if ""&request.QueryString("icono") = "1" then%>
		parent.frames[1].f.icono.value =  1
	<%end if%>
	var	ancho_tmp
	var alto_tmp

	function refrescar(dimen) {
		var alerta = false
		if (""+dimen == "inicio") {
			ancho_tmp = f.ancho_foto.value = img.width
			alto_tmp = f.alto_foto.value = img.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_foto.value = Math.round((f.alto_foto.value * img.width) / img.height)
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  ancho_tmp
				f.alto_foto.value = alto_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_foto.value = Math.round((f.ancho_foto.value * img.height) / img.width)
		}
		img.width =  f.ancho_foto.value
		img.height = f.alto_foto.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_foto.value =  f.ancho_foto.value
		parent.frames[1].f.alto_foto.value = f.alto_foto.value
		
		// Temporales para la siguiente
		ancho_tmp =  f.ancho_foto.value
		alto_tmp =  f.alto_foto.value
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}
	}

	<%if ""&request.QueryString("icono") = "auto" then%>
	// Icono
	function refrescarIcono(dimen) {
		var alerta = false
		if (""+dimen == "inicio") {
			ancho_icono_tmp = f.ancho_icono.value = img_icono.width
			alto_icono_tmp = f.alto_icono.value = img_icono.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_icono.value = Math.round((f.alto_icono.value * img_icono.width) / img_icono.height)
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  ancho_icono_tmp
				f.alto_icono.value = alto_icono_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_icono.value = Math.round((f.ancho_icono.value * img_icono.height) / img_icono.width)
		}
		img_icono.width =  f.ancho_icono.value
		img_icono.height = f.alto_icono.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_icono.value =  f.ancho_icono.value
		parent.frames[1].f.alto_icono.value = f.alto_icono.value
		
		// Temporales para la siguiente
		ancho_icono_tmp =  f.ancho_icono.value
		alto_icono_tmp =  f.alto_icono.value
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}
	}
	<%end if%>	
	
	setTimeout("refrescar('inicio');<%if ""&request.QueryString("icono") = "auto" then%>refrescarIcono('inicio');<%end if%>",250)
</script>



<%
case "opciones_foto"

	archivo = ""&request.Form("archivo")
	tipo = lcase(""&getExtension(archivo))
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Opciones de la imagen</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<script>
function numerico(c){
	if (isNaN(c.value)){
		alert("Escriba un valor numérico.")
		c.focus()
	}
}
</script>
<form name="f" method="post" action="">
<input type="hidden" name="anchomax" value="<%=anchomax_cuerpo_sitio%>">

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
	
	<%if tipo = "gif" or tipo = "jpg" then%>
		<table  border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
		  <tr>
			<td align="center"><img src="<%=archivo%>" name="id_img" id="id_img"></td>
		  </tr>
		</table>
	<%else%>
		<table width="80%">
		<tr>
		  <td bgcolor="#FFDDDD"><b>AVISO: </b><br>
		    Las im&aacute;genes subidas es necesario que sean de tipo jpg o gif.</td>
		</tr>
</table>
	<%end if%>

      <table width="80%"  border="0" align="center" cellpadding="2" cellspacing="2">
        
        <%if tipo = "jpg" then%>
		<tr align="left" valign="top" bgcolor="#FFFFFF">
          <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
                <td><font color="#0066CC">Tama&ntilde;o foto (p&iacute;xeles):</font></td>
                <td align="right"><b>Ancho</b>:
                  <input name="ancho_foto" type="text" class="campoAdmin" id="ancho_foto" size="5" value="">
                  <a href="JavaScript:refrescar('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>  <b>Alto</b>:
                  <input name="alto_foto" type="text" class="campoAdmin" id="alto_foto" size="5" value="">
                  <a href="JavaScript:refrescar('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a></td>
                <td align="right">&nbsp;</td>
              </tr>
          </table></td>
        </tr>
		<%end if%>
        
        
		<%if ""&request.QueryString("icono") = "auto" then
			if tipo <> "jpg" then%>
			<tr>
			  <td bgcolor="#FFFFEA"><b>NOTA: </b>Para generar un icono automáticamente y poder modificar los tama&ntilde;os de las im&aacute;genes subidas, es necesario que sean de tipo jpg.</td>
			</tr>
			<%else%>
		<tr valign="top">
          <td align="center" valign="middle"><table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#849ACE">
            <tr>
              <td><table width="100%"  border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                <tr>
                  <td align="center"><table  border="0" cellspacing="2" cellpadding="3">
                      <tr align="left">
                        <td colspan="2"><font color="#0066CC">Tama&ntilde;o icono (p&iacute;xeles):</font></td>
                      </tr>
                      <tr>
                        <td align="right"><b>Ancho</b>: </td>
                        <td bgcolor="#F5F5F5">&nbsp;
                            <input name="ancho_icono" type="text" class="campoAdmin" id="ancho_icono" onBlur="numerico(this)" size="5">
                            <a href="JavaScript:refrescarIcono('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>&nbsp;</td>
                      </tr>
                      <tr>
                        <td align="right"><b>Alto</b>:</td>
                        <td bgcolor="#F5F5F5">&nbsp;
                            <input name="alto_icono" type="text" class="campoAdmin" id="alto_icono" onBlur="numerico(this)" size="5">
                            <a href="JavaScript:refrescarIcono('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>&nbsp;</td>
                      </tr>
                  </table></td>
                  <td align="center"><table  border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
                      <tr>
                        <td><img src="<%=request.form("archivo")%>" name="id_img_icono" width="100" hspace="3" vspace="3" id="id_img_icono"></td>
                      </tr>
                  </table></td>
                </tr>
              </table></td>
            </tr>
          </table>
            </td>
          </tr>
			  <%end if%>
		  <%end if%>
		  
        
      </table>
	  </td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
    <%if tipo = "gif" or tipo = "jpg" then%>
	<input name="enviar" type="button" class="botonAdmin" id="enviar" onClick="refrescar('enviar');" value="Enviar">
	<%end if%>
	</td>
  </tr>
</table>
</form>
<script language="javascript" type="text/javascript">
	// Temp Globales
	img = document.getElementById("id_img")
	<%if ""&request.QueryString("icono") = "auto" then%>
		img_icono = document.getElementById("id_img_icono")
		var	ancho_icono_tmp
		var alto_icono_tmp
	<%end if%>
	<%if ""&request.QueryString("icono") = "1" then%>
		parent.frames[1].f.icono.value =  1
	<%end if%>
	var	ancho_tmp
	var alto_tmp

	function refrescar(dimen) {
		var alerta = false
	<%if tipo <> "jpg" then%>
	<%else%>
		if (""+dimen == "inicio") {
			ancho_tmp = f.ancho_foto.value = img.width
			alto_tmp = f.alto_foto.value = img.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_foto.value = Math.round((f.alto_foto.value * img.width) / img.height)
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  ancho_tmp
				f.alto_foto.value = alto_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_foto.value = Math.round((f.ancho_foto.value * img.height) / img.width)
		}
		img.width =  f.ancho_foto.value
		img.height = f.alto_foto.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_foto.value =  f.ancho_foto.value
		parent.frames[1].f.alto_foto.value = f.alto_foto.value
		
		// Temporales para la siguiente
		ancho_tmp =  f.ancho_foto.value
		alto_tmp =  f.alto_foto.value

	<%end if%>
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}

	}

	<%if ""&request.QueryString("icono") = "auto" then%>
	// Icono
	function refrescarIcono(dimen) {
		<%if tipo<>"jpg" then%>
		<%else%>
		var alerta = false
		if (""+dimen == "inicio") {
			ancho_icono_tmp = f.ancho_icono.value = img_icono.width
			alto_icono_tmp = f.alto_icono.value = img_icono.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_icono.value = Math.round((f.alto_icono.value * img_icono.width) / img_icono.height)
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  ancho_icono_tmp
				f.alto_icono.value = alto_icono_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_icono.value = Math.round((f.ancho_icono.value * img_icono.height) / img_icono.width)
		}
		img_icono.width =  f.ancho_icono.value
		img_icono.height = f.alto_icono.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_icono.value =  f.ancho_icono.value
		parent.frames[1].f.alto_icono.value = f.alto_icono.value
		
		// Temporales para la siguiente
		ancho_icono_tmp =  f.ancho_icono.value
		alto_icono_tmp =  f.alto_icono.value
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}
		<%end if%>
	}
	<%end if%>	
	

	setTimeout("refrescar('inicio');<%if ""&request.QueryString("icono") = "auto" then%>refrescarIcono('inicio');<%end if%>",250)
</script>

<%case "guardarfoto"

	id = ""&request.QueryString("id")
	if id <> "" then

		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.SaveToMemory()
		set foto = Upload.files("archivo")
		ancho_real = numero(foto.imageWidth)
		foto.SaveAs Server.MapPath(rutaDatos &"/fotos/foto"& id &".jpg")
		ancho = numero(Upload.Form("ancho_foto"))
		alto = numero(Upload.Form("alto_foto"))
		tipo = ""&lcase(foto.ImageType)
		If tipo = "jpg" Then

			Set jpeg = Server.CreateObject("Persits.Jpeg")
			jpeg.Open(foto.Path)

			if ancho > 0 and alto > 0 then

				jpeg.Width = ancho
				jpeg.Height = alto
	
				jpeg.Save server.MapPath(rutaDatos &"/fotos/foto"& id &"."& tipo)
	
			end if

			icono = Upload.Form("icono") ' Añadir icono distindo por separado.
			ancho_icono = Upload.Form("ancho_icono")
			alto_icono = Upload.Form("alto_icono")
			if ancho_icono <> "" and alto_icono <> "" then
				jpeg.Width = ancho_icono
				jpeg.Height = alto_icono
				jpeg.Save Server.MapPath(rutaDatos &"/iconos/icono"& id &"."& tipo)
			end if

		elseif tipo = "gif" then
			foto.SaveAs Server.MapPath(rutaDatos &"/fotos/foto"& id &".gif")
		else
			' La imagen debe ser gif o jpg
			unerror = true : msgerror = "La imagen debe ser gif o jpg"
		end if
		
		' MDB
		'------
		if not unerror then
			sql = "UPDATE REGISTROS SET "
			sql = sql & "R_FOTO = 'foto"&id&"."&tipo&"'"
			if ancho_icono <> "" and alto_icono <> "" then
				sql = sql & ",R_ICONO = 'icono"&id&"."&tipo&"'"
			end if
			sql = sql & " WHERE R_ID = " & id
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing
		end if
		
		if unerror then%>
			<script language="javascript" type="text/javascript">
				parent.location.href = 'inicio.asp?msgerror=<%=msgerror%>'
			</script>
		<%else%>
			<script language="javascript" type="text/javascript">
				//try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					<%if ""&request.QueryString("icono") = "1" or ""&request.QueryString("icono") = "anadir" or ""&request.QueryString("icono") = "cambiar" or ""&icono = "1" then%>
						 parent.location.href='archivos_frames.asp?ac=formguardaricono&id=<%=id%>'
					<%elseif ""&request.QueryString("icono") = "quitar" then%>
						 parent.location.href='archivos_frames.asp?ac=quitaricono&id=<%=id%>'
					<%else%>
						 parent.location.href='inicio.asp'
					<%end if%>
				//}catch(unerror){}
			</script>	
		<%
		end if
	end if

case "opciones_fotoseccion2"


	archivo = ""&request.Form("archivo")
	tipo = lcase(""&getExtension(archivo))
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Opciones de la imagen sub secci&oacute;n </font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<script>
function numerico(c){
	if (isNaN(c.value)){
		alert("Escriba un valor numérico.")
		c.focus()
	}
}
</script>
<form name="f" method="post" action="">
<input type="hidden" name="anchomax" value="<%=anchomax_cuerpo_sitio%>">

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
	
	<%if tipo = "gif" or tipo = "jpg" then%>
		<table  border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
		  <tr>
			<td align="center"><img src="<%=archivo%>" name="id_img" id="id_img"></td>
		  </tr>
		</table>
	<%else%>
		<table width="80%">
		<tr>
		  <td bgcolor="#FFDDDD"><b>AVISO: </b><br>
		    Las im&aacute;genes subidas es necesario que sean de tipo jpg o gif.</td>
		</tr>
</table>
	<%end if%>

      <table width="80%"  border="0" align="center" cellpadding="2" cellspacing="2">
        
        <%if tipo = "jpg" then%>
		<tr align="left" valign="top" bgcolor="#FFFFFF">
          <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
                <td><font color="#0066CC">Tama&ntilde;o foto (p&iacute;xeles):</font></td>
                <td align="right"><b>Ancho</b>:
                  <input name="ancho_foto" type="text" class="campoAdmin" id="ancho_foto" size="5" value="">
                  <a href="JavaScript:refrescar('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>  <b>Alto</b>:
                  <input name="alto_foto" type="text" class="campoAdmin" id="alto_foto" size="5" value="">
                  <a href="JavaScript:refrescar('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a></td>
                <td align="right">&nbsp;</td>
              </tr>
          </table></td>
        </tr>
		<%end if%>
        
        
		<%if ""&request.QueryString("icono") = "auto" then
			if tipo <> "jpg" then%>
			<tr>
			  <td bgcolor="#FFFFEA"><b>NOTA: </b>Para generar un icono automáticamente y poder modificar los tama&ntilde;os de las im&aacute;genes subidas, es necesario que sean de tipo jpg.</td>
			</tr>
			<%else%>
		<tr valign="top">
          <td align="center" valign="middle"><table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#849ACE">
            <tr>
              <td><table width="100%"  border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                <tr>
                  <td align="center"><table  border="0" cellspacing="2" cellpadding="3">
                      <tr align="left">
                        <td colspan="2"><font color="#0066CC">Tama&ntilde;o icono (p&iacute;xeles):</font></td>
                      </tr>
                      <tr>
                        <td align="right"><b>Ancho</b>: </td>
                        <td bgcolor="#F5F5F5">&nbsp;
                            <input name="ancho_icono" type="text" class="campoAdmin" id="ancho_icono" onBlur="numerico(this)" size="5">
                            <a href="JavaScript:refrescarIcono('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>&nbsp;</td>
                      </tr>
                      <tr>
                        <td align="right"><b>Alto</b>:</td>
                        <td bgcolor="#F5F5F5">&nbsp;
                            <input name="alto_icono" type="text" class="campoAdmin" id="alto_icono" onBlur="numerico(this)" size="5">
                            <a href="JavaScript:refrescarIcono('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>&nbsp;</td>
                      </tr>
                  </table></td>
                  <td align="center"><table  border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
                      <tr>
                        <td><img src="<%=request.form("archivo")%>" name="id_img_icono" width="100" hspace="3" vspace="3" id="id_img_icono"></td>
                      </tr>
                  </table></td>
                </tr>
              </table></td>
            </tr>
          </table>
            </td>
          </tr>
			  <%end if%>
		  <%end if%>
		  
        
      </table>
	  </td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
    <%if tipo = "gif" or tipo = "jpg" then%>
	<input name="enviar" type="button" class="botonAdmin" id="enviar" onClick="refrescar('enviar');" value="Enviar">
	<%end if%>
	</td>
  </tr>
</table>
</form>
<script language="javascript" type="text/javascript">
	// Temp Globales
	img = document.getElementById("id_img")
	<%if ""&request.QueryString("icono") = "auto" then%>
		img_icono = document.getElementById("id_img_icono")
		var	ancho_icono_tmp
		var alto_icono_tmp
	<%end if%>
	<%if ""&request.QueryString("icono") = "1" then%>
		parent.frames[1].f.icono.value =  1
	<%end if%>
	var	ancho_tmp
	var alto_tmp

	function refrescar(dimen) {
		var alerta = false
	<%if tipo <> "jpg" then%>
	<%else%>
		if (""+dimen == "inicio") {
			ancho_tmp = f.ancho_foto.value = img.width
			alto_tmp = f.alto_foto.value = img.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_foto.value = Math.round((f.alto_foto.value * img.width) / img.height)
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  ancho_tmp
				f.alto_foto.value = alto_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_foto.value = Math.round((f.ancho_foto.value * img.height) / img.width)
		}
		img.width =  f.ancho_foto.value
		img.height = f.alto_foto.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_foto.value =  f.ancho_foto.value
		parent.frames[1].f.alto_foto.value = f.alto_foto.value
		
		// Temporales para la siguiente
		ancho_tmp =  f.ancho_foto.value
		alto_tmp =  f.alto_foto.value

	<%end if%>
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}

	}

	<%if ""&request.QueryString("icono") = "auto" then%>
	// Icono
	function refrescarIcono(dimen) {
		<%if tipo<>"jpg" then%>
		<%else%>
		var alerta = false
		if (""+dimen == "inicio") {
			ancho_icono_tmp = f.ancho_icono.value = img_icono.width
			alto_icono_tmp = f.alto_icono.value = img_icono.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_icono.value = Math.round((f.alto_icono.value * img_icono.width) / img_icono.height)
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  ancho_icono_tmp
				f.alto_icono.value = alto_icono_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_icono.value = Math.round((f.ancho_icono.value * img_icono.height) / img_icono.width)
		}
		img_icono.width =  f.ancho_icono.value
		img_icono.height = f.alto_icono.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_icono.value =  f.ancho_icono.value
		parent.frames[1].f.alto_icono.value = f.alto_icono.value
		
		// Temporales para la siguiente
		ancho_icono_tmp =  f.ancho_icono.value
		alto_icono_tmp =  f.alto_icono.value
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}
		<%end if%>
	}
	<%end if%>	
	

	setTimeout("refrescar('inicio');<%if ""&request.QueryString("icono") = "auto" then%>refrescarIcono('inicio');<%end if%>",250)
</script>

<%case "opciones_fotoseccion"


	archivo = ""&request.Form("archivo")
	tipo = lcase(""&getExtension(archivo))
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Opciones de la imagen secci&oacute;n </font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<script>
function numerico(c){
	if (isNaN(c.value)){
		alert("Escriba un valor numérico.")
		c.focus()
	}
}
</script>
<form name="f" method="post" action="">
<input type="hidden" name="anchomax" value="<%=anchomax_cuerpo_sitio%>">

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
	
	<%if tipo = "gif" or tipo = "jpg" then%>
		<table  border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
		  <tr>
			<td align="center"><img src="<%=archivo%>" name="id_img" id="id_img"></td>
		  </tr>
		</table>
	<%else%>
		<table width="80%">
		<tr>
		  <td bgcolor="#FFDDDD"><b>AVISO: </b><br>
		    Las im&aacute;genes subidas es necesario que sean de tipo jpg o gif.</td>
		</tr>
</table>
	<%end if%>

      <table width="80%"  border="0" align="center" cellpadding="2" cellspacing="2">
        
        <%if tipo = "jpg" then%>
		<tr align="left" valign="top" bgcolor="#FFFFFF">
          <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
                <td><font color="#0066CC">Tama&ntilde;o foto (p&iacute;xeles):</font></td>
                <td align="right"><b>Ancho</b>:
                  <input name="ancho_foto" type="text" class="campoAdmin" id="ancho_foto" size="5" value="">
                  <a href="JavaScript:refrescar('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>  <b>Alto</b>:
                  <input name="alto_foto" type="text" class="campoAdmin" id="alto_foto" size="5" value="">
                  <a href="JavaScript:refrescar('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a></td>
                <td align="right">&nbsp;</td>
              </tr>
          </table></td>
        </tr>
		<%end if%>
        
        
		<%if ""&request.QueryString("icono") = "auto" then
			if tipo <> "jpg" then%>
			<tr>
			  <td bgcolor="#FFFFEA"><b>NOTA: </b>Para generar un icono automáticamente y poder modificar los tama&ntilde;os de las im&aacute;genes subidas, es necesario que sean de tipo jpg.</td>
			</tr>
			<%else%>
		<tr valign="top">
          <td align="center" valign="middle"><table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#849ACE">
            <tr>
              <td><table width="100%"  border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                <tr>
                  <td align="center"><table  border="0" cellspacing="2" cellpadding="3">
                      <tr align="left">
                        <td colspan="2"><font color="#0066CC">Tama&ntilde;o icono (p&iacute;xeles):</font></td>
                      </tr>
                      <tr>
                        <td align="right"><b>Ancho</b>: </td>
                        <td bgcolor="#F5F5F5">&nbsp;
                            <input name="ancho_icono" type="text" class="campoAdmin" id="ancho_icono" onBlur="numerico(this)" size="5">
                            <a href="JavaScript:refrescarIcono('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>&nbsp;</td>
                      </tr>
                      <tr>
                        <td align="right"><b>Alto</b>:</td>
                        <td bgcolor="#F5F5F5">&nbsp;
                            <input name="alto_icono" type="text" class="campoAdmin" id="alto_icono" onBlur="numerico(this)" size="5">
                            <a href="JavaScript:refrescarIcono('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>&nbsp;</td>
                      </tr>
                  </table></td>
                  <td align="center"><table  border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
                      <tr>
                        <td><img src="<%=request.form("archivo")%>" name="id_img_icono" width="100" hspace="3" vspace="3" id="id_img_icono"></td>
                      </tr>
                  </table></td>
                </tr>
              </table></td>
            </tr>
          </table>
            </td>
          </tr>
			  <%end if%>
		  <%end if%>
		  
        
      </table>
	  </td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
    <%if tipo = "gif" or tipo = "jpg" then%>
	<input name="enviar" type="button" class="botonAdmin" id="enviar" onClick="refrescar('enviar');" value="Enviar">
	<%end if%>
	</td>
  </tr>
</table>
</form>
<script language="javascript" type="text/javascript">
	// Temp Globales
	img = document.getElementById("id_img")
	<%if ""&request.QueryString("icono") = "auto" then%>
		img_icono = document.getElementById("id_img_icono")
		var	ancho_icono_tmp
		var alto_icono_tmp
	<%end if%>
	<%if ""&request.QueryString("icono") = "1" then%>
		parent.frames[1].f.icono.value =  1
	<%end if%>
	var	ancho_tmp
	var alto_tmp

	function refrescar(dimen) {
		var alerta = false
	<%if tipo <> "jpg" then%>
	<%else%>
		if (""+dimen == "inicio") {
			ancho_tmp = f.ancho_foto.value = img.width
			alto_tmp = f.alto_foto.value = img.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_foto.value = Math.round((f.alto_foto.value * img.width) / img.height)
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  ancho_tmp
				f.alto_foto.value = alto_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_foto.value = Math.round((f.ancho_foto.value * img.height) / img.width)
		}
		img.width =  f.ancho_foto.value
		img.height = f.alto_foto.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_foto.value =  f.ancho_foto.value
		parent.frames[1].f.alto_foto.value = f.alto_foto.value
		
		// Temporales para la siguiente
		ancho_tmp =  f.ancho_foto.value
		alto_tmp =  f.alto_foto.value

	<%end if%>
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}

	}

	<%if ""&request.QueryString("icono") = "auto" then%>
	// Icono
	function refrescarIcono(dimen) {
		<%if tipo<>"jpg" then%>
		<%else%>
		var alerta = false
		if (""+dimen == "inicio") {
			ancho_icono_tmp = f.ancho_icono.value = img_icono.width
			alto_icono_tmp = f.alto_icono.value = img_icono.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_icono.value = Math.round((f.alto_icono.value * img_icono.width) / img_icono.height)
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  ancho_icono_tmp
				f.alto_icono.value = alto_icono_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_icono.value = Math.round((f.ancho_icono.value * img_icono.height) / img_icono.width)
		}
		img_icono.width =  f.ancho_icono.value
		img_icono.height = f.alto_icono.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_icono.value =  f.ancho_icono.value
		parent.frames[1].f.alto_icono.value = f.alto_icono.value
		
		// Temporales para la siguiente
		ancho_icono_tmp =  f.ancho_icono.value
		alto_icono_tmp =  f.alto_icono.value
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}
		<%end if%>
	}
	<%end if%>	
	

	setTimeout("refrescar('inicio');<%if ""&request.QueryString("icono") = "auto" then%>refrescarIcono('inicio');<%end if%>",250)
</script>

<%case "guardarfotoseccion"


	id = ""&request.QueryString("id")
	if id <> "" then

		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.SaveToMemory()
		set foto = Upload.files("archivo")
		tipo = ""&lcase(foto.ImageType)
		foto.SaveAs Server.MapPath(rutaDatos &"/fotos_seccion/foto"& id &"."& tipo)
		ancho = numero(Upload.Form("ancho_foto"))
		alto = numero(Upload.Form("alto_foto"))
		If tipo = "jpg" Then

			Set jpeg = Server.CreateObject("Persits.Jpeg")
			jpeg.Open(foto.Path)

			if ancho > 0 and alto > 0 then
				jpeg.Width = ancho
				jpeg.Height = alto
			end if
			jpeg.Save server.MapPath(rutaDatos &"/fotos_seccion/foto"& id &"."& tipo)

			icono = Upload.Form("icono") ' Añadir icono distindo por separado.
			ancho_icono = Upload.Form("ancho_icono")
			alto_icono = Upload.Form("alto_icono")
			if ancho_icono <> "" and alto_icono <> "" then
				jpeg.Width = ancho_icono
				jpeg.Height = alto_icono
				jpeg.Save Server.MapPath(rutaDatos &"/iconos_seccion/icono"& id &"."& tipo)
			end if

		elseif tipo = "gif" then
			foto.SaveAs Server.MapPath(rutaDatos &"/fotos_seccion/foto"& id &".gif")
		else
			' La imagen debe ser gif o jpg
			unerror = true : msgerror = "La imagen debe ser gif o jpg"
		end if
		
		' MDB
		'------
		if not unerror then
			sql = "UPDATE SECCIONES SET "
			sql = sql & "S_FOTO = 'foto"&id&"."&tipo&"'"
			if ancho_icono <> "" and alto_icono <> "" then
				sql = sql & ",S_ICONO = 'icono"&id&"."&tipo&"'"
			end if
			sql = sql & " WHERE S_ID = " & id
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing
		end if
		
		if unerror then%>
			<script language="javascript" type="text/javascript">
				parent.location.href = 'inicio.asp?msgerror=<%=msgerror%>'
			</script>
		<%else%>
			<script language="javascript" type="text/javascript">
				//try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					<%if ""&request.QueryString("icono") = "1" or ""&request.QueryString("icono") = "anadir" or ""&request.QueryString("icono") = "cambiar" or ""&icono = "1" then%>
						 parent.location.href = 'archivos_frames.asp?ac=formguardaricono&id=<%=id%>'
					<%elseif ""&request.QueryString("icono") = "quitar" then%>
						 parent.location.href = 'archivos_frames.asp?ac=quitaricono&id=<%=id%>'
					<%else%>
						 parent.location.href = 'inicio.asp'
					<%end if%>
				//}catch(unerror){}
			</script>	
		<%
		end if
	end if

case "guardarfotoseccion2"


	id = ""&request.QueryString("id")
	if id <> "" then

		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.SaveToMemory()
		set foto = Upload.files("archivo")
		tipo = ""&lcase(foto.ImageType)
		foto.SaveAs Server.MapPath(rutaDatos &"/fotos_seccion2/foto"& id &"."& tipo)
		ancho = numero(Upload.Form("ancho_foto"))
		alto = numero(Upload.Form("alto_foto"))
		If tipo = "jpg" Then

			Set jpeg = Server.CreateObject("Persits.Jpeg")
			jpeg.Open(foto.Path)

			if ancho > 0 and alto > 0 then
				jpeg.Width = ancho
				jpeg.Height = alto
			end if
			jpeg.Save server.MapPath(rutaDatos &"/fotos_seccion2/foto"& id &"."& tipo)

			icono = Upload.Form("icono") ' Añadir icono distindo por separado.
			ancho_icono = Upload.Form("ancho_icono")
			alto_icono = Upload.Form("alto_icono")
			if ancho_icono <> "" and alto_icono <> "" then
				jpeg.Width = ancho_icono
				jpeg.Height = alto_icono
				jpeg.Save Server.MapPath(rutaDatos &"/iconos_seccion2/icono"& id &"."& tipo)
			end if

		elseif tipo = "gif" then
			foto.SaveAs Server.MapPath(rutaDatos &"/fotos_seccion2/foto"& id &".gif")
		else
			' La imagen debe ser gif o jpg
			unerror = true : msgerror = "La imagen debe ser gif o jpg"
		end if
		
		' MDB
		'------
		if not unerror then
			sql = "UPDATE SECCIONES2 SET "
			sql = sql & "S2_FOTO = 'foto"&id&"."&tipo&"'"
			if ancho_icono <> "" and alto_icono <> "" then
				sql = sql & ",S2_ICONO = 'icono"&id&"."&tipo&"'"
			end if
			sql = sql & " WHERE S2_ID = " & id
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing
		end if
		
		if unerror then%>
			<script language="javascript" type="text/javascript">
				parent.location.href = 'inicio.asp?msgerror=<%=msgerror%>'
			</script>
		<%else%>
			<script language="javascript" type="text/javascript">
				//try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					<%if ""&request.QueryString("icono") = "1" or ""&request.QueryString("icono") = "anadir" or ""&request.QueryString("icono") = "cambiar" or ""&icono = "1" then%>
						 parent.location.href = 'archivos_frames.asp?ac=formguardaricono2&id=<%=id%>'
					<%elseif ""&request.QueryString("icono") = "quitar" then%>
						 parent.location.href = 'archivos_frames.asp?ac=quitaricono2&id=<%=id%>'
					<%else%>
						 parent.location.href = 'inicio.asp'
					<%end if%>
				//}catch(unerror){}
			</script>	
		<%
		end if
	end if

case "guardararchivo"

	id = ""&request.QueryString("id")
	if id <> "" then
		set up = new xelUpload
		up.Upload()
		ruta = server.MapPath(rutaDatos&"/archivos")
		Response.Flush	
		For each fich in up.Ficheros.Items
		
			' Comprobaciones
			
			' Mayor de 1 MB 
			if fich.Tamano > maximo then
				unerror = true : msgerror = "El tamaño de su archivo supera el máximo permitido.<br>Su archivo: "& fich.Tamano & " - " & FormatNumber(fich.Tamano / (1024*1024)) & " Mb.<br>Máximo: "& FormatNumber(maximo / (1024*1024)) &" Mb."
			end if
			
			nombre = fich.Nombre
			extensionOriginal = right(nombre,len(nombre)-inStrRev(nombre,"."))
			
'			' Comprobar tipo
'			t = fich.TipoContenido
'			if t = "image/x-png" then
'				ex = "png"
'			elseif t = "image/pjpeg" then
'				ex = "jpg"
'			elseif t = "image/gif" then
'				ex = "gif"
'			elseif t = "audio/mpeg" then
'				ex = "mp3"
'			elseif t = "text/html" then
'				ex = "html"
'			elseif t = "application/x-zip-compressed" then
'				ex = "zip"
'			elseif t = "audio/wav" then
'				ex = "wav"
'			elseif t = "application/msword" then
'				ex = "doc"
'			elseif t = "application/msaccess" then
'				ex = "mdb"
'			else
'				unerror = true : msgerror = "No es el tipo permitido."
'			end if

		ex = extensionOriginal
			
		
			if not unerror then
				'fich.Guardar ruta
				nombre = "archivo"&id&"."&ex
				fich.GuardarComo nombre, ruta
			end if

			sql = "UPDATE REGISTROS SET"
			sql = sql & " R_ARCHIVO = '"& nombre &"',"
			sql = sql & " R_TIPOARCHIVO = '"& extensionOriginal &"'"
			sql = sql & " WHERE R_ID = " & id
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing
			

		Next
		set up = nothing
		
		if unerror then
		%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          archivo</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><b>Ha ocurrido el siguiente error:</b></td>
  </tr>
  <tr>
    <td><%=msgerror%></td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar"></td>
  </tr>
</table>

		<%else	
		%>
		<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					location.href = 'inicio.asp'
				}catch(unerror){}
		</script>	
		<%
		end if
	end if
	

case "opciones_icono"

	tipo = lcase(""&getExtension(request.Form("archivo")))
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Opciones del icono</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<script>
function numerico(c){
	if (isNaN(c.value)){
		alert("Escriba un valor numérico.")
		c.focus()
	}
}
</script>
<form name="f" method="post" action="">
<input type="hidden" name="anchomax" value="<%=icono_anchomax%>">
<%
archivo = request.Form("archivo")
tipo = lcase(""&getExtension(archivo))
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
	
	<%if tipo = "gif" or tipo = "jpg" then%>
		<table  border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
		  <tr>
			<td align="center"><img src="<%=archivo%>" name="id_img" id="id_img"></td>
		  </tr>
		</table>
	<%else%>
		<table width="80%">
		<tr>
		  <td bgcolor="#FFDDDD"><b>AVISO: </b><br>
		    Es necesario que las im&aacute;genes subidas sean de tipo jpg o gif.</td>
		</tr>
</table>
	<%end if%>

      <table width="80%"  border="0" align="center" cellpadding="2" cellspacing="2">
        
        <%if tipo = "jpg" then%>
		<tr align="left" valign="top" bgcolor="#FFFFFF">
          <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
                <td><font color="#0066CC">Tama&ntilde;o foto (p&iacute;xeles):</font></td>
                <td align="right"><b>Ancho</b>:
                  <input name="ancho_foto" type="text" class="campoAdmin" id="ancho_foto" size="5" value="">
                  <a href="JavaScript:refrescar('ancho')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>  <b>Alto</b>:
                  <input name="alto_foto" type="text" class="campoAdmin" id="alto_foto" size="5" value="">
                  <a href="JavaScript:refrescar('alto')"> <img src="/admin/global/img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a></td>
                <td align="right">&nbsp;</td>
              </tr>
          </table></td>
        </tr>
		<%end if%>
        
        
		  
        
      </table>
	  </td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
    <%if tipo = "gif" or tipo = "jpg" then%>
	<input name="enviar" type="button" class="botonAdmin" id="enviar" onClick="refrescar('enviar');" value="Enviar">
	<%end if%>
	</td>
  </tr>
</table>
</form>
<script language="javascript" type="text/javascript">
	// Temp Globales
	img = document.getElementById("id_img")
	<%if ""&request.QueryString("icono") = "auto" then%>
		img_icono = document.getElementById("id_img_icono")
		var	ancho_icono_tmp
		var alto_icono_tmp
	<%end if%>
	<%if ""&request.QueryString("icono") = "1" then%>
		parent.frames[1].f.icono.value =  1
	<%end if%>
	var	ancho_tmp
	var alto_tmp

	function refrescar(dimen) {
		var alerta = false
	<%if tipo <> "jpg" then%>
	<%else%>
		if (""+dimen == "inicio") {
			ancho_tmp = f.ancho_foto.value = img.width
			alto_tmp = f.alto_foto.value = img.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_foto.value = Math.round((f.alto_foto.value * img.width) / img.height)
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  ancho_tmp
				f.alto_foto.value = alto_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_foto.value) > Number(f.anchomax.value)){
				f.ancho_foto.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_foto.value = Math.round((f.ancho_foto.value * img.height) / img.width)
		}
		img.width =  f.ancho_foto.value
		img.height = f.alto_foto.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_foto.value =  f.ancho_foto.value
		parent.frames[1].f.alto_foto.value = f.alto_foto.value
		
		// Temporales para la siguiente
		ancho_tmp =  f.ancho_foto.value
		alto_tmp =  f.alto_foto.value

	<%end if%>
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}

	}

	<%if ""&request.QueryString("icono") = "auto" then%>
	// Icono
	function refrescarIcono(dimen) {
		<%if tipo<>"jpg" then%>
		<%else%>
		var alerta = false
		if (""+dimen == "inicio") {
			ancho_icono_tmp = f.ancho_icono.value = img_icono.width
			alto_icono_tmp = f.alto_icono.value = img_icono.height
		}
		if (""+dimen == "alto") {
			// Rd3
			f.ancho_icono.value = Math.round((f.alto_icono.value * img_icono.width) / img_icono.height)
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  ancho_icono_tmp
				f.alto_icono.value = alto_icono_tmp
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
		}else {
			if (Number(f.ancho_icono.value) > Number(f.anchomax.value)){
				f.ancho_icono.value =  f.anchomax.value
				alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
				alerta = true
			}
			// Rd3
			f.alto_icono.value = Math.round((f.ancho_icono.value * img_icono.height) / img_icono.width)
		}
		img_icono.width =  f.ancho_icono.value
		img_icono.height = f.alto_icono.value
		
		// Pasar datos al form oculto
		parent.frames[1].f.ancho_icono.value =  f.ancho_icono.value
		parent.frames[1].f.alto_icono.value = f.alto_icono.value
		
		// Temporales para la siguiente
		ancho_icono_tmp =  f.ancho_icono.value
		alto_icono_tmp =  f.alto_icono.value
		
		// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
		if (""+dimen == 'enviar' && !alerta){
			parent.frames[1].f.enviar.click()
		}
		<%end if%>
	}
	<%end if%>	
	

	setTimeout("refrescar('inicio');<%if ""&request.QueryString("icono") = "auto" then%>refrescarIcono('inicio');<%end if%>",250)
</script>

<%case "guardaricono"

	Response.Write "<b>guardaricono</b>"

	id = ""&request.QueryString("id")
	if id <> "" then

		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.SaveToMemory()
		set foto = Upload.files("archivo")
		tipo = ""&lcase(foto.ImageType)
		foto.SaveAs Server.MapPath(rutaDatos &"/iconos/icono"& id &"." & tipo)
		ancho = numero(Upload.Form("ancho_foto"))
		alto = numero(Upload.Form("alto_foto"))
		If tipo = "jpg" Then

			Set jpeg = Server.CreateObject("Persits.Jpeg")
			jpeg.Open(foto.Path)

			if ancho > 0 and alto > 0 then
				jpeg.Width = ancho
				jpeg.Height = alto
			end if
			jpeg.Save server.MapPath(rutaDatos &"/iconos/icono"& id &"."& tipo)

		elseif tipo = "gif" then
			foto.SaveAs Server.MapPath(rutaDatos &"/iconos/icono"& id &".gif")
		else
			' La imagen debe ser gif o jpg
			unerror = true : msgerror = "La imagen debe ser gif o jpg"
		end if
		
		' MDB
		'------
		if not unerror then
			sql = "UPDATE REGISTROS SET "
			sql = sql & "R_ICONO = 'icono"&id&"."&tipo&"'"
			sql = sql & " WHERE R_ID = " & id
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing
		end if
		
		if unerror then%>
			<script language="javascript" type="text/javascript">
				parent.location.href = 'inicio.asp?msgerror=<%=msgerror%>'
			</script>
		<%else%>
			<script language="javascript" type="text/javascript">
				//try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
/					f.submit()
					<%if ""&request.QueryString("icono") = "1" or ""&request.QueryString("icono") = "anadir" or ""&request.QueryString("icono") = "cambiar" or ""&icono = "1" then%>
						 parent.location.href = 'archivos.asp?ac=formguardaricono&id=<%=id%>'
					<%else%>
						 parent.location.href = 'inicio.asp'
					<%end if%>
				//}catch(unerror){}
			</script>	
		<%
		end if
	end if
	
case "guardariconoseccion"

	id = ""&request.QueryString("id")
	if id <> "" then
		set up = new xelUpload
		up.Upload()
		ruta = server.MapPath(rutaDatos&"/iconos_seccion")
		Response.Flush	
		For each fich in up.Ficheros.Items
		
			' Comprobaciones
			
			' Mayor de 1 MB 
			if fich.Tamano > maximo then
				unerror = true : msgerror = "El tamaño de su archivo supera el máximo permitido.<br>Su archivo: "& fich.Tamano & " - " & FormatNumber(fich.Tamano / (1024*1024)) & " Mb.<br>Máximo: "& FormatNumber(maximo / (1024*1024)) &" Mb."
			end if
			
			' Comprobar tipo
			t = fich.TipoContenido
			if t = "image/x-png" then
				ex = "png"
			elseif t = "image/pjpeg" then
				ex = "jpg"
			elseif t = "image/gif" then
				ex = "gif"
			else
				unerror = true : msgerror = "No es el tipo permitido."
			end if
			
			if not unerror then
				'fich.Guardar ruta
				nombre = "icono"&id&"."&ex
				fich.GuardarComo nombre, ruta
			end if

			sql = "UPDATE SECCIONES SET "
			sql = sql & "S_ICONO = '"& nombre &"'"
			sql = sql & " WHERE S_ID = " & id
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing
			

		Next
		set up = nothing

		if unerror then
		%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          icono a sección</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><b>Ha ocurrido el siguiente error:</b></td>
  </tr>
  <tr>
    <td><%=msgerror%></td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar"></td>
  </tr>
</table>

		<%else	
		%>
		<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()

					location.href = 'main.asp?ac=adminsecciones&seccion=<%=request.QueryString("seccion")%>'
				}catch(unerror){
	//
				}
		</script>	
		<%
		end if
	end if

case "guardariconoseccion2"

	id = ""&request.QueryString("id")
	if id <> "" then
		set up = new xelUpload
		up.Upload()
		ruta = server.MapPath(rutaDatos&"/iconos_seccion2")
		Response.Flush	
		For each fich in up.Ficheros.Items
		
			' Comprobaciones
			
			' Mayor de 1 MB 
			if fich.Tamano > maximo then
				unerror = true : msgerror = "El tamaño de su archivo supera el máximo permitido.<br>Su archivo: "& fich.Tamano & " - " & FormatNumber(fich.Tamano / (1024*1024)) & " Mb.<br>Máximo: "& FormatNumber(maximo / (1024*1024)) &" Mb."
			end if
			
			' Comprobar tipo
			t = fich.TipoContenido
			if t = "image/x-png" then
				ex = "png"
			elseif t = "image/pjpeg" then
				ex = "jpg"
			elseif t = "image/gif" then
				ex = "gif"
			else
				unerror = true : msgerror = "No es el tipo permitido."
			end if
			
			if not unerror then
				'fich.Guardar ruta
				nombre = "icono"&id&"."&ex
				fich.GuardarComo nombre, ruta
			end if

			sql = "UPDATE SECCIONES2 SET "
			sql = sql & "S2_ICONO = '"& nombre &"'"
			sql = sql & " WHERE S2_ID = " & id
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing
			

		Next
		set up = nothing

		if unerror then
		%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          icono a subsección</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><b>Ha ocurrido el siguiente error:</b></td>
  </tr>
  <tr>
    <td><%=msgerror%></td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar"></td>
  </tr>
</table>

		<%else	
		%>
		<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()

					location.href = 'main.asp?ac=adminsecciones2&seccion=<%=request.QueryString("seccion")%>'
				}catch(unerror){
	//
				}
		</script>	
		<%
		end if
	end if
	
case "quitararchivo"
	
			' Borrar foto de la BD
			sql = "SELECT R_ARCHIVO, R_ID FROM REGISTROS "
			sql = sql & " WHERE R_ID = " & request.QueryString("id")
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conn_
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			archivo = re("R_ARCHIVO")
			re("R_ARCHIVO") = ""
			re.update
			set re = nothing
			
			' Borrar el archivo del disco duro
			call borrararchivo(server.MapPath(rutaDatos&"/archivos/"&archivo))
			%>
					<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					location.href = 'inicio.asp'
				}catch(unerror){}
		</script>
		<%
	
case "quitarfoto"
	
			' Borrar foto de la BD
			sql = "SELECT R_FOTO, R_ID FROM REGISTROS "
			sql = sql & " WHERE R_ID = " & request.QueryString("id")
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conn_
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			foto = re("R_FOTO")
			re("R_FOTO") = ""
			re.update
			set re = nothing
			
			' Borrar foto del disco duro
			call borrararchivo(server.MapPath(rutaDatos&"/fotos/"&foto))

			select case request.QueryString("icono")
			case "quitar"
				Response.Redirect("archivos_frames.asp?ac=quitaricono&id="& request.QueryString("id") &"&archivo="& request.QueryString("archivo"))
			case "anadir"
				Response.Redirect("archivos_frames.asp?ac=formguardaricono&id="& request.QueryString("id") &"&archivo="& request.QueryString("archivo"))
			case "cambiar"
				Response.Redirect("archivos_frames.asp?ac=formguardaricono&id="& request.QueryString("id") &"&archivo="& request.QueryString("archivo"))
			case else%>
				<script>
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					location.href = 'inicio.asp'
				</script>
			<%end select

case "quitarfotoseccion"
	
			' Borrar foto de la BD
			sql = "SELECT S_FOTO, S_ID FROM SECCIONES "
			sql = sql & " WHERE S_ID = " & request.QueryString("id")
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conn_
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			foto = re("S_FOTO")
			re("S_FOTO") = ""
			re.update
			set re = nothing
			
			' Borrar foto del disco duro
			call borrararchivo(server.MapPath(rutaDatos&"/fotos_seccion/"&foto))
			%>
					<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					<%if request("icono") = "anadir" or request("icono") = "cambiar" then%>
						location.href = 'archivos.asp?ac=formguardariconoseccion&id=<%=request.QueryString("id")%>'
					<%elseif request("icono")="eliminar" then%>
						location.href = 'archivos.asp?ac=quitariconoseccion&id=<%=request.QueryString("id")%>'
					<%else%>
					location.href = 'main.asp?ac=adminsecciones&seccion=<%=request.QueryString("seccion")%>'
					<%end if%>
				}catch(unerror){}
		</script>
		<%

case "quitarfotoseccion2"
	
			' Borrar foto de la BD
			sql = "SELECT S2_FOTO, S2_ID FROM SECCIONES2 "
			sql = sql & " WHERE S2_ID = " & request.QueryString("id")
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conn_
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			foto = re("S2_FOTO")
			re("S2_FOTO") = ""
			re.update
			set re = nothing
			
			' Borrar foto del disco duro
			call borrararchivo(server.MapPath(rutaDatos&"/fotos_seccion2/"&foto))
			%>
					<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					<%if request("icono") = "anadir" or request("icono") = "cambiar" then%>
						location.href = 'archivos.asp?ac=formguardariconoseccion2&id=<%=request.QueryString("id")%>&seccion=<%=request.QueryString("seccion")%>'
					<%elseif request("icono")="eliminar" then%>
						location.href = 'archivos.asp?ac=quitariconoseccion2&id=<%=request.QueryString("id")%>&seccion=<%=request.QueryString("seccion")%>'
					<%else%>
					location.href = 'main.asp?ac=adminsecciones2&seccion=<%=request.QueryString("seccion")%>'
					<%end if%>
				}catch(unerror){}
		</script>
		<%

case "quitaricono"
	
			' Borrar icono de la BD
			sql = "SELECT R_ICONO, R_ID FROM REGISTROS "
			sql = sql & " WHERE R_ID = " & request.QueryString("id")
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conn_
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			icono = re("R_ICONO")
			re("R_ICONO") = ""
			re.update
			set re = nothing
			
			' Borrar icono del disco duro
			call borrararchivo(server.MapPath(rutaDatos&"/iconos/"&icono))
			%>
					<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					location.href = 'inicio.asp'
				}catch(unerror){}

		</script>
		<%

case "quitariconoseccion"
	
			' Borrar icono de la BD
			sql = "SELECT S_ICONO, S_ID FROM SECCIONES "
			sql = sql & " WHERE S_ID = " & request.QueryString("id")
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conn_
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			icono = re("S_ICONO")
			re("S_ICONO") = ""
			re.update
			set re = nothing
			
			' Borrar icono del disco duro
			call borrararchivo(server.MapPath(rutaDatos&"/iconos_seccion/"&icono))
			%>
					<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					location.href = 'main.asp?ac=adminsecciones&seccion=<%=request.QueryString("seccion")%>'
				}catch(unerror){}
		</script>
		<%

case "quitariconoseccion2"
	
			' Borrar icono de la BD
			sql = "SELECT S2_ICONO, S2_ID FROM SECCIONES2 "
			sql = sql & " WHERE S2_ID = " & request.QueryString("id")
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conn_
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			icono = re("S2_ICONO")
			re("S2_ICONO") = ""
			re.update
			set re = nothing
			
			' Borrar icono del disco duro
			call borrararchivo(server.MapPath(rutaDatos&"/iconos_seccion2/"&icono))
			%>
			<script>
				try{
					var f = top.frames[1].frames[0].f // Frame de la izquierda
					f.ac.value = ""
					f.action = "main.asp"
					f.target = ""
					f.submit()
					location.href = 'main.asp?ac=adminsecciones2&seccion=<%=request.QueryString("seccion")%>'
				}catch(unerror){}
		</script>
<%


case "formguardarfoto"

	if not config_iexplorer then
		%><!--#include file="inc_archivos_formguardarfoto2.asp" --><%
	else
		%><!--#include file="inc_archivos_formguardarfoto.asp" --><%
	end if

case "formguardarfotoseccion2"


%>
<script>
	function envio(){
		if (""+ f.archivo.value == ""){
			alert("Seleccione una foto.");
			return false;
		}
		f.enviar.disabled = true
		return true
	}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          imagen sub secci&oacute;n </font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<form name="f" method="post" action="archivos.asp?ac=opciones_fotoseccion2&id=<%=id%>&icono=<%=request.QueryString("icono")%>" onSubmit="return envio()">
  <table width="100%"  border="0" cellspacing="0" cellpadding="2">
  
  
  
  <tr>
    <td colspan="2" align="center" valign="middle">    
<table  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><input name="archivo" type="text" class="campoAdmin" id="archivo" size="50" readonly="true"></td>
          <td>&nbsp;
            <input type="button" class="botonAdmin" onClick="parent.frames[1].f.archivo.click()" value="Examinar ..."></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td colspan="2" align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><table  border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFCC">
      <tr>
        <td><b>Formatos v&aacute;lidos</b>: jpg o gif (s&oacute;lo formato RGB)</td>
      </tr>
    </table></td>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
      <input name="enviar" type="submit" class="botonAdmin" id="enviar" value="Enviar"></td>
  </tr>
</table>
</form>

<%case "formguardarfotoseccion"


%>
<script>
	function envio(){
		if (""+ f.archivo.value == ""){
			alert("Seleccione una foto.");
			return false;
		}
		f.enviar.disabled = true
		return true
	}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          imagen secci&oacute;n </font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<form name="f" method="post" action="archivos.asp?ac=opciones_fotoseccion&id=<%=id%>&icono=<%=request.QueryString("icono")%>" onSubmit="return envio()">
  <table width="100%"  border="0" cellspacing="0" cellpadding="2">
  
  
  
  <tr>
    <td colspan="2" align="center" valign="middle">    
<table  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><input name="archivo" type="text" class="campoAdmin" id="archivo" size="50" readonly="true"></td>
          <td>&nbsp;
            <input type="button" class="botonAdmin" onClick="parent.frames[1].f.archivo.click()" value="Examinar ..."></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td colspan="2" align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><table  border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFCC">
      <tr>
        <td><b>Formatos v&aacute;lidos</b>: jpg o gif (s&oacute;lo formato RGB)</td>
      </tr>
    </table></td>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
      <input name="enviar" type="submit" class="botonAdmin" id="enviar" value="Enviar"></td>
  </tr>
</table>
</form>


<%case "formguardararchivo"

	id = request.QueryString("id")
	if ""&id <> "" and isNumeric(id) then
%>
        <br>
<script>
	function envio(){
//		var tablaprin = parent.frames[0].tablaprin // Frame de la izquierda
//		tablaprin.disabled = true
		f.enviar.disabled = true
		return true
	}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          archivos</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<form name="f" enctype="multipart/form-data" method="post" action="archivos.asp?ac=guardararchivo&id=<%=id%>&icono=<%=request.QueryString("icono")%>" onSubmit="return envio()">
<table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td>Escoja un archivo</td>
  </tr>
  <tr>
    <td>
      <input name="archivo" type="file" class="campoAdmin" style="width:100%">
    </td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
    <input name="enviar" type="submit" class="botonAdmin" id="enviar" value="Enviar"></td>
  </tr>
</table>
</form>
<%end if

case "formguardaricono"



	id = numero(request.QueryString("id"))
	if id > 0 then

	' Abrir el registro...
	sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
	consultaXOpen sql,1	

	

%>
<script>
	function envio(){
		if (""+ f.archivo.value == ""){
			alert("Seleccione una foto.");
			return false;
		}
		f.enviar.disabled = true
		return true
	}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          icono</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<form name="f" method="post" action="archivos.asp?ac=opciones_icono&id=<%=id%>" onSubmit="return envio()">
  <table width="100%"  border="0" cellspacing="0" cellpadding="2">


	<tr>
    <td colspan="2" align="center" valign="middle">
	
	<%if reTotal >0 then%>
		<table width="100%"  border="0" cellpadding="2" cellspacing="0">
			<tr>
			  <td><b><%=config_nom_titulo%>: </b></td>
		  </tr>
			<tr>
			<td bgcolor="#FFFFFF"><%=re("R_TITULO")%></td>
			</tr>
		</table>
		<br>
	<%end if%>
	
	
      <br>
      <table  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><input name="archivo" type="text" class="campoAdmin" id="archivo" size="50" readonly="true"></td>
          <td>&nbsp;
            <input type="button" class="botonAdmin" onClick="parent.frames[1].f.archivo.click()" value="Examinar ..."></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td colspan="2" align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><table  border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFCC">
      <tr>
        <td><b>Formatos v&aacute;lidos</b>: jpg o gif (s&oacute;lo formato RGB)</td>
      </tr>
    </table></td>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
      <input name="enviar" type="submit" class="botonAdmin" id="enviar" value="Enviar"></td>
  </tr>
</table>
</form>
<%
	consultaXClose()
end if

case "formguardariconoseccion"

	id = request.QueryString("id")
	if esNumero(id) then
%>
        <br>
<script>
	function envio(){
//		var tablaprin = parent.frames[0].tablaprin // Frame de la izquierda
//		tablaprin.disabled = true
		f.enviar.disabled = true
		return true
	}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          icono a sección</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<form name="f" enctype="multipart/form-data" method="post" action="archivos.asp?ac=guardariconoseccion&id=<%=id%>&icono=<%=request.QueryString("icono")%>" onSubmit="return envio()">
<table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td>Escoja un archivo</td>
  </tr>
  <tr>
    <td>
      <input name="archivo" type="file" class="campoAdmin" style="width:100%">
    </td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
    <input name="enviar" type="submit" class="botonAdmin" id="enviar" value="Enviar"></td>
  </tr>
</table>
</form>
<%end if

case "formguardariconoseccion2"

	id = request.QueryString("id")
	if esNumero(id) then
%>
        <br>
<script>
	function envio(){
//		var tablaprin = parent.frames[0].tablaprin // Frame de la izquierda
//		tablaprin.disabled = true
		f.enviar.disabled = true
		return true
	}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar
          icono a subsección</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<form name="f" enctype="multipart/form-data" method="post" action="archivos.asp?ac=guardariconoseccion2&id=<%=id%>&icono=<%=request.QueryString("icono")%>&seccion=<%=request.QueryString("seccion")%>" onSubmit="return envio()">
<table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td>Escoja un archivo</td>
  </tr>
  <tr>
    <td>
      <input name="archivo" type="file" class="campoAdmin" style="width:100%">
    </td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="right"><input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
    <input name="enviar" type="submit" class="botonAdmin" id="enviar" value="Enviar"></td>
  </tr>
</table>
</form>
<%end if

case "ampliarfoto"%>

<%foto = "/datos/"& session("idioma") &"/"& session("cualid") &"/fotos/"& request.QueryString("archivo")
if existe(server.MapPath(foto)) then%>
	<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
	  <tr>
	    <td align="center" valign="middle"><img src="<%=foto%>" onLoad="tamano(this.width,this.height)"></td>
	  </tr>
	</table>
<%else

' Borrar el nombre de la foto del registro para que no moles
			sql = "UPDATE REGISTROS SET "
			sql = sql & "R_FOTO = ''"
			sql = sql & " WHERE R_FOTO = '"& request.QueryString("archivo") &"'"
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing


%>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><b>La foto no est&aacute; en
    el servidor</b><br>
    Se ha desvinculado del registro </td>
  </tr>
</table>
<script>
tamano(250,100)
</script>
<%end if%>

<%case "ampliarfotoseccion"%>
<%foto = "/datos/"& session("idioma") &"/"& session("cualid") &"/fotos_seccion/"& request.QueryString("archivo")
if existe(server.MapPath(foto)) then%>
	<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
	  <tr>
	    <td align="center" valign="middle"><img src="<%=foto%>" onLoad="tamano(this.width,this.height)"></td>
	  </tr>
	</table>
<%else

' Borrar el nombre de la foto del registro para que no moles
			sql = "UPDATE SECCIONES SET "
			sql = sql & "S_FOTO = ''"
			sql = sql & " WHERE S_FOTO = '"& request.QueryString("archivo") &"'"
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing


%>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><b>La foto no est&aacute; en
    el servidor</b><br>
    Se ha desvinculado del registro </td>
  </tr>
</table>
<script>
tamano(250,100)
</script>
<%end if

case "ampliarfotoseccion2"%>
<%foto = "../../datos/"& session("idioma") &"/"& session("cualid") &"/fotos_seccion2/"& request.QueryString("archivo")
if existe(server.MapPath(foto)) then%>
	<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
	  <tr>
	    <td align="center" valign="middle"><img src="<%=foto%>" onLoad="tamano(this.width,this.height)"></td>
	  </tr>
	</table>
<%else

' Borrar el nombre de la foto del registro para que no moles
			sql = "UPDATE SECCIONES SET "
			sql = sql & "S_FOTO = ''"
			sql = sql & " WHERE S_FOTO = '"& request.QueryString("archivo") &"'"
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing


%>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><b>La foto no est&aacute; en
    el servidor</b><br>
    Se ha desvinculado del registro </td>
  </tr>
</table>
<script>
tamano(250,100)
</script>
<%end if
case "ampliaricono"%>
<%icono = "../../datos/"& session("idioma") &"/"& session("cualid") &"/iconos/"& request.QueryString("archivo")
if existe(server.MapPath(icono)) then%>
	<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
	  <tr>
	    <td align="center" valign="middle"><img src="<%=icono%>" onLoad="tamano(this.width,this.height)"></td>
	  </tr>
	</table>
<%else

' Borrar el nombre del icono del registro
			sql = "UPDATE REGISTROS SET "
			sql = sql & "R_ICONO = ''"
			sql = sql & " WHERE R_ICONO = '"& request.QueryString("archivo") &"'"
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing


%>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><b>El icono no est&aacute; en
    el servidor</b><br>
    Se ha desvinculado del registro </td>
  </tr>
</table>
<script>
tamano(250,100)
</script>
<%end if

case "ampliariconoseccion"%>
<%icono = "/" & c_s & "datos/"& session("idioma") &"/"& session("cualid") &"/iconos_seccion/"& request.QueryString("archivo")
if existe(server.MapPath(icono)) then%>
	<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
	  <tr>
	    <td align="center" valign="middle"><img src="<%=icono%>" onLoad="tamano(this.width,this.height)"></td>
	  </tr>
	</table>
<%else

' Borrar el nombre del icono del registro
			sql = "UPDATE SECCIONES SET "
			sql = sql & "S_ICONO = ''"
			sql = sql & " WHERE S_ICONO = '"& request.QueryString("archivo") &"'"
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing


%>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><b>El icono no est&aacute; en
    el servidor</b><br>
    Se ha desvinculado del registro </td>
  </tr>
</table>
<script>
tamano(250,100)
</script>
<%end if

case "ampliariconoseccion2"%>
<%icono = "/" & c_s & "datos/"& session("idioma") &"/"& session("cualid") &"/iconos_seccion2/"& request.QueryString("archivo")
if existe(server.MapPath(icono)) then%>
	<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
	  <tr>
	    <td align="center" valign="middle"><img src="<%=icono%>" onLoad="tamano(this.width,this.height)"></td>
	  </tr>
	</table>
<%else

' Borrar el nombre del icono del registro
			sql = "UPDATE SECCIONES2 SET "
			sql = sql & "S2_ICONO = ''"
			sql = sql & " WHERE S2_ICONO = '"& request.QueryString("archivo") &"'"
			set oConn = server.CreateObject("ADODB.Connection")
			oConn.Open conn_
			oConn.execute sql
			oConn.Close
			set oConn = nothing


%>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><b>El icono no est&aacute; en
    el servidor</b><br>
    Se ha desvinculado del registro </td>
  </tr>
</table>
<script>
tamano(250,100)
</script>
<%end if


case else

response.Write("No se especificó una acción")

end select
%>
</body>
</html>
