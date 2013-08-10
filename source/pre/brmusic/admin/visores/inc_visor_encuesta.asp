<%
	unerror_encuesta = false
	msgerror_encuesta = ""
	ruta_xml_encuesta = "/" & c_s & "datos/" & idioma & "/encuestas/encuestas.xml"

	if idioma = "" then
		unerror_encuesta = true : msgerror_encuesta = "No se ha definido un idioma para la encuesta."
	end if
	
	if not unerror_encuesta then
		Set xml_encuesta = CreateObject("MSXML.DOMDocument")
		if not xml_encuesta.Load(server.MapPath(ruta_xml_encuesta)) then
			unerror_encuesta = true : msgerror_encuesta = "No se ha podido cargar el XML*<br>["& ruta_xml_encuesta &"]."
		else
			set datos = xml_encuesta.selectSingleNode("datos")
			if not typeOK(datos) then
				unerror_encuesta = true : msgerror_encuesta = "Error en XML de datos."
			else
				set nodo_encuesta = datos.selectSingleNode("dato")
				if not typeOK(nodo_encuesta) then
					unerror_encuesta = true : msgerror_encuesta = "No hay encuestas"
				else
					set nodo_titulo = nodo_encuesta.selectSingleNode("titulo")
					if typeOK(nodo_titulo) then
						titulo = ""& nodo_titulo.text
					end if
					set nodo_cuestion = nodo_encuesta.selectSingleNode("cuestion")
					if typeOK(nodo_cuestion) then
						cuestion = ""& nodo_cuestion.text
					end if
					if cuestion = "" or titulo = "" then
						unerror_encuesta = true : msgerror_encuesta = "Los datos de esta encuesta no estan completos."
					end if
				end if
			end if
		end if
	end if

	if not unerror_encuesta then%>
		<style type="text/css">
		<!--
		.radio {
			height: 11px;
			width: 11px;
		}
		.botonVotar {
			color: #FFFFFF;
			background-color: #457445;
			font-weight: bold;
			height: 14px;
			border: none;
			font-family: Verdana, Arial, Helvetica, sans-serif;
			font-size: 7.5pt;
		}
		-->
		</style>
		<form name="f_encuesta" method="post" action="index.asp?secc=/encuestas">
		<input type="hidden" name="id_encuesta" value="<%=nodo_encuesta.getAttribute("id")%>">
		<table width="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td colspan="2"><b><%=uncode(titulo)%></b></td>
          </tr>
          <tr>
            <td colspan="2">
			
			<table>
			<%for each opcion in nodo_encuesta.childNodes
				if inStr(opcion.nodeName,"opcion")>0 and ""&opcion.text <> ""then%>
					<tr>
					<td colspan="2">
					<table width="100%"  border="0" cellspacing="0" cellpadding="2">
						<tr>
						<td valign="top">
						<table width="100%"  border="0" cellspacing="0" cellpadding="0">
							<tr>
							<td height="2"><img src="../../spacer.gif" width="1" height="1"></td>
							</tr>
							<tr>
							<td><input name="voto" type="radio" class="radio" id="encu_<%=opcion.nodeName%>" value="<%=opcion.nodeName%>"></td>
							</tr>
						</table></td>
						<td width="100%" valign="top"><font size="1"><label for="encu_<%=opcion.nodeName%>"><%=uncode(opcion.text)%></label></font></td>
						</tr>
					</table></td>
					</tr>
				<%end if
			next%>
			</table>
			
			</td>
          </tr>
          <tr valign="bottom">
            <td><a href="index.asp?secc=/encuestas&id_encuesta=<%=nodo_encuesta.getAttribute("id")%>"><font size="1">Saber más </font></a></td>
            <td align="right"><input name="imageField" type="image" src="/sd/esp/imagenes/votar.gif" width="58" height="17" border="0"></td>
          </tr>
        </table>
		
		</form>
	<%end if
	
	if unerror_encuesta then
		if session("usuario") = 1 then
			Response.Write "<b>Error en Encuesta:</b> "& msgerror_encuesta
		end if
	end if
%>
