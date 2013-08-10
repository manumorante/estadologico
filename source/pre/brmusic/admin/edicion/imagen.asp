<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/admin/inc_rutinas.asp" -->
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<!--#include virtual="/admin/global/xelupload.asp" -->
<!--#include file="inc_swfheader.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Imagen</title>
</head>
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" class="bodyAdmin">
<%

Dim idi
Dim num
Dim archivoXml

idi = ""&request("idi")
secc = ""&request("secc")
num = ""&request("num")

if idi = "" or secc = "" or num = "" then
	unerror = true : msgerror = "No se han recibido todos los parámetros necesarios.Idi: "&idi&" | Secc: "& secc &" | Num: "& num &"."
end if

if not unerror then
	if not getPermisoPara("edicion", idi, session("usuario")) then
		unerror = true : msgerror = "Usted no tiene derechos de aSkipper para esta zona."
	end if
end if

if not unerror then
	archivoXml = "/"& c_s & idi & secc & "/" &nombrearchivo(secc) & ".xml"
	dim xmlObj
	set xmlObj = CreateObject("MSXML.DOMDocument")
	if not xmlObj.Load(Server.MapPath(archivoXml)) then
		unerror = true : msgerror = "No se ha encontrado el archivo que desea editar."
	end if
end if

if not unerror then
	dim nodoImagen
	set nodoImagen = xmlObj.selectSingleNode("contenido/imagen"&num)
	if not typeOK(nodoImagen) then
		unerror = true : msgerror = "No se ha encontrado 'imagen"&num&"' en el XML."
	end if	
end if

select case request.QueryString("ac")
case "previo"

	if secc = "" then
		unerror = true : msgerror = "No se ha recibido la SECC."
	end if

	haynuevafoto = false
	if not unerror then
		if "" & request.form("fotoNueva") <> "" then
			haynuevafoto = true
			' Condiciones requeridas en el XML
			anchoMax = 0+ ("0"&request.form("anchomax"))
			altoMax = 0+ ("0"&request.form("altomax"))
			anchoMin = 0+ ("0"&request.form("anchomin"))
			altoMin =0+ ("0"& request.form("altomin"))
			
			if ""&anchoMax = "" then anchoMax = 600 end if
			if ""&altoMax = "" then altoMax = 800 end if
			if ""&anchoMin = "" then anchoMin = 0 end if
			if ""&altoMin = "" then altoMin = 0 end if
			
			' Atributos del archivo subido
			nodo = ""&request.form("nodo")
			img = ""&request.form("fotoNueva")
			tipo = getExtension(img)
			Originalpath = ""&request.form("fotoNueva")
			
			
			if tipo = "SWF" then
				' Propiedades de Flash
				set myObj = new swfdump
				myObj.SWFDump (server.MapPath("/flah.swf"))
				alto = myObj.Heigt
				ancho = myObj.Width
				myObj.Version
			else
				if tipo <> "GIF" and tipo <> "JPG" then
					error_img = true : msgerror_img = "El archivo introducido es de tipo <b>"& tipo &"</b>.<br>Por favor, Escoja una imagen de tipo <b>GIF</b> o <b>JPG</b>."
				end if
			end if
		end if  ' img <> ""
	end if

		%>
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20" bgcolor="#FFFFFF"><span class="tituloazonaAdmin">Nueva imagen</span></td>
  </tr>
</table>
		<br>
		
		
		<form action="#" method="post" enctype="multipart/form-data" name="f">
			<input type="hidden" name="num" value="<%=num%>">
			<input type="hidden" name="secc" value="<%=secc%>">
			<input type="hidden" name="idi" value="<%=idi%>">
			<input type="hidden" name="img" value="<%=img%>">
			<input type="hidden" name="imgtemp" value="<%=imgtemp%>">
			<input type="hidden" name="tipo" value="<%=tipo%>">
			<input type="hidden" name="altoant" value="<%=request.form("altoant")%>">			
			<input type="hidden" name="anchoant" value="<%=request.form("anchoant")%>">			
			<input type="hidden" name="comentario_imagen" value="<%=request.form("comentario_imagen")%>">			
			<input type="hidden" name="enlaceventana" value="<%=request.form("enlaceventana")%>">			
			<input type="hidden" name="nodo" value="<%=request.form("nodo")%>">			
			<input type="hidden" name="pie" value="<%=request.form("pie")%>">			
			<input type="hidden" name="anchomin" value="<%=request.form("anchomin")%>">			
			<input type="hidden" name="anchomax" value="<%=request.form("anchomax")%>">			
			<input type="hidden" name="altomax" value="<%=request.form("altomax")%>">			
			<input type="hidden" name="altomin" value="<%=request.form("altomin")%>">			
			<input type="hidden" name="enlace" value="<%=request.form("enlace")%>">			
			<input type="hidden" name="fotoAnterior" value="<%=request.form("fotoAnterior")%>">			
			<input type="hidden" name="fotoNueva" value="<%=request.form("fotoNueva")%>">
			
			<%if not haynuevafoto then%>
			<script>
				top.frames["imagen_form"].f.enviar.click()
			</script>
			<%else%>
		<table  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#7996B0">
          <tr>
            <td colspan="2"><table  border="0" align="center" cellpadding="4" cellspacing="0">

			<%if not error_img then%>
              <tr>
                <td align="center" bgcolor="#FFFFFF">
				<table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td align="left">&nbsp;<font size="+1">Vista previa</font></td>
                  </tr>
                  <tr>
                    <td align="center">
					<%if tipo = "SWF" then%>
					<script language="javascript" type="text/javascript">
					<!--
						var ancho_ori
						var alto_ori
						function dimensiones(ancho,alto){
							if (ancho > f.anchomax.value) {
								alto = Math.round((alto * f.anchomax.value) / ancho)
								ancho = f.anchomax.value
							}
							ancho_ori = ancho
							alto_ori = alto
							f.ancho.value = ancho
							f.alto.value = alto
							top.frames[1].f.ancho.value = ancho
							top.frames[1].f.alto.value = alto
						}
					//-->
					</script>
					  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="653" height="484">
                          <param name="movie" value="flash.swf?archivo=<%=Originalpath%>">
                          <param name="quality" value="high">
                          <embed src="flash.swf?archivo=<%=Originalpath%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="653" height="484"></embed>
					  </object>
					<%else%>
						
						<img id="id_img" src="<%=Originalpath%>">
					<%end if%>
					</td>
                  </tr>
                </table>
				</td>
              </tr>
              <%end if%>
              <tr>
                <td bgcolor="#F0F0F0">
				<table  border="0" align="center" cellpadding="2" cellspacing="2">
                    <tr valign="top">
                      <td align="right"><b>Archivo</b>: </td>
                      <td colspan="2" bgcolor="#FFFFFF">&nbsp;<%=replace(Originalpath,img,"<b>"&img&"</b>")%>&nbsp;</td>
                  </tr>
		       <tr align="left" valign="top">
				     <td height="4" colspan="3"><img src="../../spacer.gif" width="1" height="1"></td>
			      </tr>
				  <% if tipo = "JPG" then %>
				   <tr align="left" valign="top">
				     <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td><font color="#0066CC">T</font><font color="#0066CC">ama&ntilde;o (p&iacute;xeles):</font></td>
                         <td align="right">&nbsp;</td>
                       </tr>
                     </table>			          </td>
		          </tr>
				
				   <tr valign="top">
                      <td align="right" valign="top"><b>Ancho</b>: </td>
                     <td valign="top" bgcolor="#FFFFFF">&nbsp;
                        <input name="ancho" type="text" class="campoAdmin" id="ancho" size="5" value="">
                     <a href="JavaScript:refrescar('ancho')"> <img src="../../img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>                     </td>
                      <td align="right" bgcolor="#FFFFFF"><font color="#000055">m&aacute;x. <%=request.form("anchomax")%> px</font></td>
                  </tr>
                    <tr valign="top">
                      <td align="right" valign="top"><b>Alto</b>: </td>
                      <td valign="top" bgcolor="#FFFFFF">&nbsp;
                        <input name="alto" type="text" class="campoAdmin" id="alto" size="5" value="">
                      <a href="JavaScript:refrescar('alto')"> <img src="../../img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a>                      </td>
                      <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>				
				  <% elseif tipo = "SWF" then %>  
				   <tr valign="top">
                      <td align="right" valign="top"><b>Ancho</b>: </td>
                     <td valign="top" bgcolor="#FFFFFF">&nbsp;
                        <input name="ancho" type="text" class="campoAdmin" id="ancho" size="5" value="">
                    </td>
                      <td align="right" bgcolor="#FFFFFF"><font color="#000055">m&aacute;x. <%=request.form("anchomax")%> px</font></td>
                  </tr>
                    <tr valign="top">
                      <td align="right" valign="top"><b>Alto</b>: </td>
                      <td valign="top" bgcolor="#FFFFFF">&nbsp;
                        <input name="alto" type="text" class="campoAdmin" id="alto" size="5" value="">
                      </td>
                      <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>	
					<% elseif tipo = "GIF" then%>
				   <tr valign="top">
                      <td align="right" valign="top"><b>Ancho</b>: </td>
                     <td valign="top" bgcolor="#FFFFFF">&nbsp;
                        <input name="ancho" type="text" class="campoAdmin" id="ancho" size="5" readonly="true">
                    </td>
                      <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                    <tr valign="top">
                      <td align="right" valign="top"><b>Alto</b>: </td>
                      <td valign="top" bgcolor="#FFFFFF">&nbsp;
                        <input name="alto" type="text" class="campoAdmin" id="alto" size="5" readonly="true">
                      </td>
                      <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>	
					<% end if %>
                </table>
<%if tipo = "SWF" then%>
	<script language="javascript" type="text/javascript">
		function refrescar(dimen){
			unerror = false
			if (isNaN(f.ancho.value)){;
				alert("El ancho debe ser de tipo numérico.");
				f.ancho.value = ancho_ori;
				f.ancho.focus();
				unerror = true;
			}
			if (isNaN(f.alto.value) && !unerror){
				alert("El alto debe ser de tipo numérico.");
				f.alto.value = alto_ori;
				f.alto.focus();
				unerror = true;
			}
			if (f.ancho.value > f.anchomax.value && !unerror){
				alert("El ancho debe ser menor o igual al ancho máximo de la página: " + f.anchomax.value +".");
				f.ancho.value = ancho_ori;
				f.ancho.focus();
				unerror = true;
			}
			if (unerror == false) {
				top.frames[1].f.ancho.value = f.ancho.value;
				top.frames[1].f.alto.value = f.alto.value;
				top.frames[1].f.enviar.click();
			}
		}
	</script>
<%elseif tipo = "JPG" then%>
	<%if not error_img then%>
		<script language="javascript" type="text/javascript">
			// Temp Globales
			img = document.getElementById("id_img")
			var	ancho_tmp
			var alto_tmp

			function refrescar(dimen) {
					
				var alerta = false
				if (""+dimen == "inicio") {
					ancho_tmp = f.ancho.value = img.width
					alto_tmp = f.alto.value = img.height
				}
				if (""+dimen == "alto") {
					// Rd3
					f.ancho.value = Math.round((f.alto.value * img.width) / img.height)
					if (Number(f.ancho.value) > Number(f.anchomax.value)){
						f.ancho.value =  ancho_tmp
						f.alto.value = alto_tmp
						alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
						alerta = true
					}
				}else {
					if (Number(f.ancho.value) > Number(f.anchomax.value)){
						f.ancho.value =  f.anchomax.value
						alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
						alerta = true
					}
					// Rd3
					f.alto.value = Math.round((f.ancho.value * img.height) / img.width)
				}
				img.width =  f.ancho.value
				img.height = f.alto.value
				
				// Pasar datos al form oculto
				top.frames[1].f.ancho.value =  f.ancho.value
				top.frames[1].f.alto.value = f.alto.value
				
				// Temporales para la siguiente
				ancho_tmp =  f.ancho.value
				alto_tmp =  f.alto.value
				
				// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
				if (""+dimen == 'enviar' && !alerta){
					top.frames[1].f.enviar.click()
					f.aceptar.disabled = true
					f.aceptar.value = "Espere ..."
					f.aceptar.title= "Por favor, espere."
				}
			}
		</script>
	<%end if%>

<%elseif tipo = "GIF" then%>
	<%if not error_img then%>
		<script language="javascript" type="text/javascript">
			// Temp Globales
			img = document.getElementById("id_img")
			var	ancho_tmp
			var alto_tmp

			function refrescar(dimen) {
				
				var alerta = false
				if (""+dimen == "inicio") {
					ancho_tmp = f.ancho.value = img.width
					alto_tmp = f.alto.value = img.height
				}
				
				// Pasar datos al form oculto
				top.frames[1].f.ancho.value =  f.ancho.value
				top.frames[1].f.alto.value = f.alto.value
				
				// Temporales para la siguiente
				ancho_tmp =  f.ancho.value
				alto_tmp =  f.alto.value
				
				// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
				if (""+dimen == 'enviar' && !alerta){
					top.frames[1].f.enviar.click()
					f.aceptar.disabled = true
					f.aceptar.value = "Espere ..."
					f.aceptar.title= "Por favor, espere."
				}
			}
		</script>
	<%end if%>

<%end if%>


				
				</td>
              </tr>

			<%if error_img then%>
              <tr>
                <td bgcolor="#FFFFFF"><font color="#FF0000"><b>ATENCI&Oacute;N:<br>
                </b><%=msgerror_img%></font> </td>
              </tr>
              <%end if%>
            </table></td>
          </tr>
		<%if not error_img then%>

          <tr>
            <td height="3" colspan="2" align="right" bgcolor="#F5F5F5"><img src="../../spacer.gif" width="1" height="1"></td>
          </tr>
          <tr>
            <td colspan="2" align="right" bgcolor="#F5F5F5">
				
				<input name="" type="button" class="botonAdmin" onClick="top.window.close()" value="Cancelar">
				<input name="aceptar" type="button" class="botonAdmin" id="aceptar" onClick="refrescar('enviar');" value="Aceptar">
			
			</td>
          </tr>

		  <%else%>
          <tr>
            <td colspan="2" align="right" bgcolor="#F5F5F5"><input name="" type="button" class="botonAdmin" onClick="location.href='<%=request.ServerVariables("HTTP_REFERER")%>'" value="Volver">
            </td>
          </tr>
		  <%end if%>

        </table>
		
		<%end if ' if haynuevafoto%>
		
</form>
<%if haynuevafoto and tipo <> "SWF" then%>
<script>
	setTimeout("refrescar('inicio');",500)
</script>
<%end if%>
		<br>

		<%

	'end if
	


case "eliminar"

	carpeta = "/" & c_s & idi & secc
	img = ""&request.QueryString("img")
	
	if img <> "" then
		call borrarArchivo(server.MapPath(carpeta &"/"& img))
		nombreicono=Left(img,len(img)-4)&"_movil.jpg"
		call borrarArchivo(server.MapPath(carpeta &"/"& nombreicono))
	end if

	nodoImagen.text = ""

	' Pie de foto
	set att = xmlObj.createAttribute("pie")
	nodoImagen.setAttributeNode(att)
	att.nodeValue = ""

	' Enlace
	set att = xmlObj.createAttribute("enlace")
	nodoImagen.setAttributeNode(att)
	att.nodeValue = ""

	' Enlace ventana
	set att = xmlObj.createAttribute("enlaceventana")
	nodoImagen.setAttributeNode(att)
	att.nodeValue = ""

	' No visible
	set att = xmlObj.createAttribute("novisible")
	nodoImagen.setAttributeNode(att)
	att.nodeValue = ""

	' Borde/margen
	set att = xmlObj.createAttribute("margen")
	nodoImagen.setAttributeNode(att)
	att.nodeValue = ""
	
	set att = nothing

	xmlObj.save Server.MapPath(archivoXml)

	if not unerror then
		%>
		<script language="javascript" type="text/javascript">
			parent.opener.location.href=parent.opener.location
			top.window.close()
		</script>
		<%
	end if


case "ok"

	if not unerror then
		secc = request.QueryString("secc")
		num = request.QueryString("num")
		idi = request.QueryString("idi")
		if secc = "" or num = "" or idi = "" then
			unerror = true : msgerror = "Faltan datos necesarios."
		end if
	end if
	
	if not unerror then
		'on error resume next
		Set Upload = Server.CreateObject("Persits.Upload")
		Path = Server.MapPath("/" & c_s & idi & secc)
		Count = Upload.Save(Path)
		If Count > 0 Then
			Set File = Upload.Files(1)
			tipo = Ucase(""&getExtension(file.OriginalPath))
			If tipo = "JPG" Then
				' Create instance of AspJpeg object
	'			on error resume next
				Set jpeg = Server.CreateObject("Persits.Jpeg")
				jpeg.Open( File.Path )
				jpeg.Width = int(numero(Upload.Form("ancho")))
				jpeg.Height = int(numero(Upload.Form("alto")))
			
				SavePath = Path & "/" & File.ExtractFileName
				

				
				' AspJpeg always generates thumbnails in JPEG format.
				' If the original file was not a JPEG, append .JPG ext.
				If UCase(Right(SavePath, 3)) <> "JPG" Then
					SavePath = SavePath & ".jpg"
				End If
				
				jpeg.Save SavePath



		'ESTA OTRA IMAGEN ES PARA LOS MÓVILES

				
				proporcion= int(numero(Upload.Form("ancho")))/176
				jpeg.Width = 176
				jpeg.Height = int(numero(Upload.Form("alto"))/proporcion)
				
				SavePath = Path & "/" & File.ExtractFileName 

				nombreicono=Left(SavePath,len(SavePath)-4)&"_movil.jpg"
			jpeg.Save nombreicono
			
							
				on error goto 0
			end if
		end if		

		if err<>0 then
			unerror = true : msgerror = err.description
		end if
		on error goto 0
	end if

	if not unerror then
		hayfotonueva = false
		set archivo =  Upload.files("archivo")
		if typeOK(archivo) then
			hayfotonueva = true
		end if
		editar = false
		if ""&Upload.form("editar") = "1" then
			editar = true
		end if
	end if

	if hayfotonueva then
		ruta = archivo.path
		nombre = nombreArchivo(ruta)
		tipo = ucase(getExtension(nombre))

		dim ancho, alto
		ancho = int(numero(Upload.Form("ancho")))
		alto = int(numero(Upload.Form("alto")))
		
		if ancho <= 0 or alto <= 0 then
			unerror = true : msgerror = "No se ha recibido el ancho o el alto del archivo"
		end if
	end if

	' Borrar la imagen que exista actualmente.
	if not editar and hayfotonueva then
		ruta_imagen_actual = ""&nodoImagen.text
		if ruta_imagen_actual <> nombre then
			if ruta_imagen_actual <> "" then
				carpeta = "/" & c_s & idi & secc
				call borrarArchivo(server.MapPath(carpeta &"/"& ruta_imagen_actual))
				nombreicono=Left(ruta_imagen_actual,len(ruta_imagen_actual)-4)&"_movil.jpg"
				call borrarArchivo(server.MapPath(carpeta &"/"& nombreicono))
			end if
		end if
	end if

	if not unerror then
		if hayfotonueva then
			nodoImagen.text = nombre
		end if

		if editar then
			tipo = ucase(getExtension(""&Upload.Form("archivo_actual")))
		end if
		
		if editar and tipo <> "SWF" then
			' Aplicar el nuevo tamaño a la foto (ASPJpeg)
			Set jpeg = Server.CreateObject("Persits.Jpeg")
			jpeg.Open(Path &"/"& Upload.Form("archivo_actual"))

			ancho = int(numero(Upload.Form("ancho")))
			alto = int(numero(Upload.Form("alto")))
			if ancho > 0 and alto > 0 then
				jpeg.Width = ancho
				jpeg.Height = alto
			end if
			jpeg.Save Path & "/" & Upload.Form("archivo_actual")
			set jpeg = nothing
		end if
		
		if hayfotonueva or editar then
			' Ancho
			set att = xmlObj.createAttribute("ancho")
			nodoImagen.setAttributeNode(att)
			att.nodeValue = int(numero(Upload.Form("ancho")))
			' Alto
			set att = xmlObj.createAttribute("alto")
			nodoImagen.setAttributeNode(att)
			att.nodeValue = int(numero(Upload.Form("alto")))
		end if
		
		' Pie de foto
		set att = xmlObj.createAttribute("pie")
		nodoImagen.setAttributeNode(att)
		att.nodeValue = Upload.Form("pie")
	
		' Enlace
		if tipodato = "swf" then
			enlace = ""
		else
			enlace = ""&Upload.Form("enlace")
		end if
		if trim(enlace) = "http://" then
			enlace = ""
		end if
		set att = xmlObj.createAttribute("enlace")
		nodoImagen.setAttributeNode(att)
		att.nodeValue = enlace
	
		' Enlace ventana
		set att = xmlObj.createAttribute("enlaceventana")
		nodoImagen.setAttributeNode(att)
		att.nodeValue = Upload.Form("enlaceventana")
	
		' No visible
		set att = xmlObj.createAttribute("novisible")
		nodoImagen.setAttributeNode(att)
		att.nodeValue = Upload.Form("novisible")
	
		' Borde/margen
		set att = xmlObj.createAttribute("margen")
		nodoImagen.setAttributeNode(att)
		att.nodeValue = Upload.Form("margen")

		set att = nothing
		xmlObj.save Server.MapPath(archivoXml)
	end if

	if not unerror then
		%>
		<script language="javascript" type="text/javascript">
			parent.opener.location.href=parent.opener.location
			top.window.close()
		</script>
		<%
	end if
	
	if unerror then
		Response.Write "<b>Error</b>: " & msgerror
	end if


case "limpiar"

	call borrarArchivo(server.MapPath(request.form("imgtemp")))
	Response.Redirect("imagen.asp?secc="& request.form("secc") &"&idi="& request.form("idi") &"&num="& request.form("num") &"")

case else

	if not unerror then%>
			<table width="100%"  border="0" cellspacing="0" cellpadding="10">
			  <tr>
				<td>
				<%
'				formImagen (n,nombreCampo,nombreFoto,ruta,anchoMax,altoMax,anchoMin,altoMin,margen,novisible,pie,enlace,enlaceventana)
'				formImagen (num,nodoImagen.nodeName,nodoImagen.text,rutavuelta,nodoImagen.getAttribute("anchomax"),nodoImagen.getAttribute("altomax"),nodoImagen.getAttribute("anchomin"),nodoImagen.getAttribute("altomin"),nodoImagen.getAttribute("margen"),nodoImagen.getAttribute("novisible"),nodoImagen.getAttribute("pie"),nodoImagen.getAttribute("enlace"),nodoImagen.getAttribute("enlaceventana"))
				nombreFoto = ""&nodoImagen.text
				'tipo = getExtension(nombreFoto)
				ancho = numero (nodoImagen.getAttribute("ancho"))
				alto = numero (nodoImagen.getAttribute("alto"))
				tmp = numero(eval("ancho_"& nodoImagen.getAttribute("anchomax") &"_cuerpo_sitio"))
				if tmp >0 then
					anchoMax = tmp
				else
					anchoMax = nodoImagen.getAttribute("anchomax")
				end if
				comentario = ""&nodoImagen.getAttribute("comentario")
				pie = ""&nodoImagen.getAttribute("pie")
				novisible = ""&nodoImagen.getAttribute("novisible")
				margen = ""&nodoImagen.getAttribute("margen")				
				enlace = ""&nodoImagen.getAttribute("enlace")
				enlaceventana = ""&nodoImagen.getAttribute("enlaceventana")
				%>
				<script language="javascript" type="text/javascript">
				<!--
					function envio() {
//						try{
							refrescarEditar("ancho")
							// Pasar datos a form oculto
							var fo = top.frames["imagen_form"].f
							fo.enlace.value = f.enlace.value
							if (f.enlaceventana_self.checked){
								fo.enlaceventana.value = "_self"
							} else {
								fo.enlaceventana.value = "_blank"
							}
							fo.pie.value = f.pie.value
							fo.novisible.checked = f.novisible.checked
							fo.margen.checked = f.margen.checked
						
							f.enviar_btn.disabled = true
							f.cancelar_btn.disabled = true
							if (f.eliminar_btn){
								f.eliminar_btn.disabled = true
							}
//						} catch(unerror){}
					}

					

				//-->
				</script>
				<form action="imagen.asp?ac=previo&idi=<%=idi%>&num=<%=num%>&secc=<%=secc%>" method="post" name="f" onSubmit="envio()">
				<input type="hidden" name="nodo" value="imagen<%=n%>">

				<input type="hidden" name="anchomax" value=<%=anchoMax%>>
				<input type="hidden" name="altomax" value=<%=altoMax%>>
				<input type="hidden" name="anchomin" value=<%=anchoMin%>>
				<input type="hidden" name="altomin" value=<%=altoMin%>>
				<table width="98%" border="0" align="center" cellpadding="2" cellspacing="0">
					<tr>
					<td align="left" valign="top">
					<%
					if nombreFoto <> "" then
						Set fso = Server.CreateObject("Scripting.FileSystemObject")
						laRutaImagen =  "../../"& session("idioma") & secc& "/" & nombreFoto
						tipo = Ucase(""&getExtension(laRutaImagen))
						if fso.FileExists(Server.MapPath(laRutaImagen)) then%>
						<input type="hidden" name="fotoactual" value="<%=laRutaImagen%>">
							<table  border="0" cellpadding="0" cellspacing="0">
								<tr>
								<td align="center" valign="middle">
								<%if tipo = "SWF" then%>
								<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="<%=ancho%>" height="<%=alto%>">
                                  <param name="movie" value="<%=laRutaImagen%>">
                                  <param name="quality" value="high">
                                  <embed src="<%=laRutaImagen%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="<%=ancho%>" height="<%=alto%>"></embed>
								  </object>
								<%else%>
								<img src="<%=laRutaImagen%>" alt="<%=nombreFoto%>" name="id_img" border="0" id="id_img">
								<%end if%>

								</td>
								</tr>
					  </table>
							<br>
						<%else%>
							<br><font color="#990000" title="La imagen indicada en el XML no se encuentra en el servidor."><nobr>Imagen no encontrada<br>en el servidor.</nobr></font>
						<%end if
					else%>
						<br>
						Introduzca una imagen o un Flash.<br>
						<br>
					<%end if%></td>
					<td width="150" align="center" valign="middle">

					<%if nombreFoto <> "" then%>
					<table  border="0" cellpadding="1" cellspacing="0">
						<tr>
						<td bgcolor="#FF0000"><table width="100%" border="0" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
						<%if numero(anchoMax) > 0 then%>
							<tr>
							<td align="right">Ancho m&aacute;x.:</td>
							<td><b><%=anchoMax%> px</b></td>
							</tr>
						<%end if%>
						</table></td>
						</tr>
						<tr>
						  <td height="4"><img src="../../spacer.gif" width="1" height="1"></td>
					  </tr>
						<tr>
						  <td bgcolor="#316AC5"><table width="100%"  border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                            <% if tipo="JPG" then %>
                            <tr align="left" valign="top">
                              <td colspan="2"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td><font color="#0066CC">Tama&ntilde;o (p&iacute;xeles):</font></td>
                                  </tr>
                              </table></td>
                            </tr>
                            <tr>
                              <td align="right"><b>Ancho</b>:</td>
                              <td align="left">&nbsp;
                                  <input name="ancho" type="text" disabled="true" class="campoAdmin" id="ancho" onBlur="refrescarEditar('ancho')" value="<%=ancho%>" size="5">
                                  <a href="JavaScript:refrescarEditar('ancho')"> <img src="../../img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a> </td>
                            </tr>
                            <tr>
                              <td align="right"><b>Alto</b>:</td>
                              <td align="left">&nbsp;
                                  <input name="alto" type="text" disabled="true" class="campoAdmin" id="alto" onBlur="refrescarEditar('alto')" value="<%=alto%>" size="5">
                                  <a href="JavaScript:refrescarEditar('alto')"> <img src="../../img/actualizar.gif" alt=" Refrescar " width="14" height="14" border="0" align="absmiddle"></a> </td>
                            </tr>
                            <% elseif tipo = "SWF" then %>
                            <tr>
                              <td align="right"><b>Ancho</b>:</td>
                              <td align="left">&nbsp;
                                  <input name="ancho" type="text" disabled="true" class="campoAdmin" id="ancho" onBlur="refrescarEditar('ancho')" value="<%=ancho%>" size="5">
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><b>Alto</b>:</td>
                              <td align="left">&nbsp;
                                  <input name="alto" type="text" disabled="true" class="campoAdmin" id="alto" onBlur="refrescarEditar('alto')" value="<%=alto%>" size="5">
                              </td>
                            </tr>
							<% elseif tipo = "GIF" then%>
                            <tr>
                              <td align="right"><b>Ancho</b>:</td>
                              <td align="left">&nbsp;
                                  <input name="ancho" type="text" class="campoAdmin" id="ancho" onBlur="refrescarEditar('ancho')" value="<%=ancho%>" size="5" readonly="true">
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><b>Alto</b>:</td>
                              <td align="left">&nbsp;
                                  <input name="alto" type="text" class="campoAdmin" id="alto" onBlur="refrescarEditar('alto')" value="<%=alto%>" size="5" readonly="true">
                              </td>
                            </tr>
                            <% end if %>
                          </table></td>
					  </tr>
					  
					  <%if tipo = "JPG" or tipo = "SWF" then%>
						<tr>
						  <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><input name="editar" type="checkbox" id="editar" onClick="refrescarEditar('ancho')" value="1"></td>
                              <td width="100%"><label for="editar">Aplicar cambios</label></td>
                            </tr>
                          </table></td>
					  </tr>
					  <%end if%>
					  </table>
					  <%end if%>
					
					  
					  </td>
				  </tr>
				  </table>
				  
				  <script language="javascript" type="text/javascript">
					<%if inStr(laRutaImagen,".swf")>0  then%>
					<%else%>
						img = document.getElementById("id_img")
					<%end if%>

					function refrescarEditar(dimen){
						<%if nombreFoto <> "" then%>
						if (f.editar){
							if (f.editar.checked){
								f.ancho.disabled = false
								f.alto.disabled = false
							} else {
								f.ancho.disabled = true
								f.alto.disabled = true
							}
						}
						<%if tipo="SWF" then%>
							if (Number(f.ancho.value) > Number(f.anchomax.value)){
								f.ancho.value =  ancho_tmp
								f.alto.value = alto_tmp
								alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
								alerta = true
							}
						<%else%>
							var alerta = false
							if (""+dimen == "alto") {
								// Rd3
								f.ancho.value = Math.round((f.alto.value * img.width) / img.height)
								if (Number(f.ancho.value) > Number(f.anchomax.value)){
									f.ancho.value =  ancho_tmp
									f.alto.value = alto_tmp
									alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
									alerta = true
								}
							}else {
								if (Number(f.ancho.value) > Number(f.anchomax.value)){
									f.ancho.value =  f.anchomax.value
									alert("El ancho máximo permitido es "+ f.anchomax.value + " px")
									alerta = true
								}
								// Rd3
								if (img){
									f.alto.value = Math.round((f.ancho.value * img.height) / img.width)
								}else{
									f.alto.value = 0
									f.ancho.value = 0
								}
							}
							if (img){
								img.width =  f.ancho.value
								img.height = f.alto.value
							}
						<%end if%>
						
						// Pasar datos al form oculto
						top.frames[1].f.ancho.value =  f.ancho.value;
						top.frames[1].f.alto.value = f.alto.value;
						<%if tipo = "JPG" or tipo = "SWF" then%>
							top.frames[1].f.editar.checked = f.editar.checked;
						<%else%>
							top.frames[1].f.editar.checked = false;
						<%end if%>
						top.frames[1].f.archivo_actual.value = "<%=nombreFoto%>";
						
						// Temporales para la siguiente
						ancho_tmp =  f.ancho.value
						alto_tmp =  f.alto.value
						
						// Si todo es correcto y se ha ordenado enviar, ENVIAMOS
						if (""+dimen == 'enviar' && !alerta){
							top.frames[1].f.enviar.click()
						}
						<%end if%>
					}
				  </script>
					<input type="hidden" name="fotoAnterior" value="<%=nombreFoto%>">
							      <fieldset>
                                  <legend>Escoja su nueva imagen</legend>
								  <table width="100%"  border="0" align="center" cellpadding="1" cellspacing="0">
                                    <tr>
                                      <td><textarea name="fotoNueva" cols="80" rows="" readonly="readonly" wrap="virtual" style="width:100%;height:26px;"></textarea></td>
                                      <td width="75">

									  <button onClick="top.frames[1].f.archivo.click()" type="button" name="" style="width:100px">
										  <table border="0" cellpadding="1" cellspacing="0" width="100%">
											  <tr valign="middle">
											    <td align="left"><img src="../images/imagen.gif" width="18" height="18" hspace="2"></td>
											  <td>Examinar... </td>
											  </tr>
										  </table>
									  </button>

									  </td>
                                    </tr>
                                    
                                  </table>
							      </fieldset>
							      <br>
							      <fieldset>
							      <legend>Enlace</legend>
							      <table width="100%" border="0" cellspacing="0" cellpadding="1">
								  <tr>
								  <td>
								  <%
								  if enlace = "" then
									  enlace = "http://"
								  end if%>
								  <input name="enlace" type="text" class="campoAdmin" id="enlace" style="width:100%" value="<%=enlace%>" maxlength="250"></td>
								  </tr>
								  <tr>
								    <td valign="middle"><b>Se abrir&aacute; en:</b>  <input name="enlaceventana" id="enlaceventana_self" type="radio" value="_self" <%if enlaceventana = "_self" or enlaceventana = "" then%>checked<%end if%>> <label for="enlaceventana_self">Misma ventana</label> <input name="enlaceventana" id="enlaceventana_blank" type="radio" value="_blank" <%if enlaceventana = "_blank" then%>checked<%end if%>> <label for="enlaceventana_blank">Ventana nueva</label></td>
								    </tr>
								  </table>
							      </fieldset>
							      <br>
							      <fieldset>
                                  <legend>Pie de foto</legend>
                                  <table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                    <tr>
                                      <td><textarea name="pie" cols="50" rows="3" wrap="virtual" class="areaAdmin" style="width:100%"><%=pie%></textarea></td>
                                    </tr>
                                  </table>
						        </fieldset>
						        <br>
						        <table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                  <tr>
                                    <td><%if novisible <> "-1" then%>
                                        <input name="novisible" type="checkbox" id="novisible" value="1" <%if novisible="1" then response.write("checked") end if%>>
                                        <label for="novisible">No visible</label>
                                        <%else%>
                                        <input type="hidden" name="novisible" value="-1">
                                        <%end if
								if margen <> "-1" then%>
                                        <input name="margen" type="checkbox" id="margen" value="1" <%if margen="1" then response.write("checked") end if%>>
                                        <label for="margen">Con borde</label>
                                        <%else%>
                                        <input type="hidden" name="margen" value="-1">
                                        <%end if%></td>
                                  </tr>
                                </table>
					            <input type="hidden" name="comentario_imagen" value="<%=comentario%>">
					            <table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
						
						<tr>
						  <td align="right"><input name="cancelar_btn" type="button" class="botonAdmin" id="cancelar_btn" onClick="top.window.close()" value="Cancelar">
                            <%if nombreFoto <> "" then%>
                            <input name="eliminar_btn" type="button" class="botonAdmin" id="eliminar_btn" title=" Eliminar esta imagen " onClick="location.href='imagen.asp?ac=eliminar&secc=<%=secc%>&img=<%=nombreFoto%>&idi=<%=idi%>&num=<%=num%>'" value="Eliminar">
                            <%end if%>
                            <input name="enviar_btn" type="submit" class="botonAdmin" id="enviar_btn" value="Enviar"></td>
					  </tr>
					</table>
				
				<input type="hidden" name="anchoant" value="<%=ancho%>">
				<input type="hidden" name="altoant" value="<%=alto%>">
				
				</form></td>


			  </tr>
</table>
	<%end if

end select

if unerror then
	Response.Write "<b>Error:</b><br>"&msgerror
end if

Function nombrearchivo(valor)
	for n=0 to len(valor)-1
		if Mid(valor,len(valor)-n,1)="\" or Mid(valor,len(valor)-n,1)="/" then
			n=len(valor)+1
		else
			nombrearchivo=Mid(valor,len(valor)-n,1)&nombrearchivo
		end if
	next
end Function

' Convierte lo que le pasemos a número, en caso de nos ser un número válido devuelve 0
		function numero(n)
			if ""&n <> "" then
				n = replace(n,".",",")
				if isNumeric(n) then
					numero = 0+n
				else
					numero = 0
				end if
			end if
			
		end function%>
</body>
</html>