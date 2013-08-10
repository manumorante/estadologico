<%Dim unerror, msgerror, rextotal, rex%>
<!--#include file="inc_conn.asp" -->
<!--#include virtual="/admin/global/inc_inicia_xml.asp" -->

<%if unerror then%>
	<b>Error</b><br><%=msgerror%>
<%else

	inicia_xml
	
	if ""&idioma = "" then
		pi = request.ServerVariables("PATH_INFO")
		if inStr(pi,"/esp/index.asp") then
			idioma = "esp"
		elseif inStr(pi,"/eng/index.asp") then
			idioma = "eng"
		elseif inStr(pi,"/fra/index.asp") then
			idioma = "fra"
		elseif inStr(pi,"/deu/index.asp") then
			idioma = "deu"
		elseif inStr(pi,"/ita/index.asp") then
			idioma = "ita"
		else
			if session("idioma") <> "" then
				idioma = session("idioma")
			else
				unerror = true : msgerror = "No ha sido posible determinar el idioma de navegación / aSkipper"
			end if
		end if
	end if
	
	

	' Ac
	ac = ""&request("ac")
	id = ""&request.QueryString("id")

	' Paginado
	registrosPorPagina = 5

	'Inicio (Cabeza lectora)
	Dim pag
	pag = ""&request("pag")
	if pag <> "" then
		if not esNumero(pag) then
			pag = 1
		end if
	else
		pag = 1
	end if

	' Sección 
	Dim seccion
	if esNumero(request("seccion")) then
		seccion = request("seccion")
	else
		seccion = -1
	end if
	
	' Sección2
	Dim seccion2
	if esNumero(request("seccion2")) then
		seccion2 = request("seccion2")
	else
		seccion2 = -1
	end if
	
	' Cadena de búsqueda
	Dim cadena
	if ""&request("cadena") <> "" then
		cadena = request("cadena")
	else
		cadena = ""
	end if



if not unerror then%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
	<td> 
	
<!-- INICIO DE FORMULARIO --------------------------------------------------------------------------------------- -->
<form name="f" action="../admin/visores/index.asp?secc=<%=secc%>" method="POST" onSubmit="return envio()">
<input type="hidden" name="ac" value="<%=ac%>">
<input type="hidden" name="ref" value="">
<input type="hidden" name="id" value="">
<input type="hidden" name="pag" value="<%=pag%>">
<input type="hidden" name="seccion" value="<%=seccion%>">
<input type="hidden" name="seccion2" value="<%=seccion2%>">
<input type="hidden" name="seccion3" value="<%=seccion3%>">

<%if nav_buscar then%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="noticias-busqueda">
		<tr>
		<td>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
			<td align="right" valign="middle">Texto a buscar: 
			  <input name="cadena" type="text" class="noticias-input" value="<%=cadena%>" size="30" maxlength="150">
			  &nbsp;
			  <input name="cadenaanterior" type="hidden" value="<%=cadena%>"></td>
			<td valign="bottom">
			<input type="image" src="/<%=c_s%><%=idioma%>/imagenes/buscar.gif">
			<a href="JavaScript:todo();"><img src="/<%=c_s%><%=idioma%>/imagenes/vertodo.gif" border="0"></a></td>
			</tr>
		</table></td>
		</tr>
	</table>
	<%else%>
	<input type="hidden" name="cadena" value="">
	<input name="cadenaanterior" type="hidden" value="<%=cadena%>">
<%end if%>

<script language="JavaScript" type="text/javascript">
	<!--
	// envio
	function envio() {
		return true
	}
	// todo
	function todo() {
		f.ac.value = ""
		f.cadena.value = ""
		f.seccion.value = ""
		f.cadenaanterior.value = ""
		f.pag.value = 1
		f.submit()
	}
	//
	// volver
	function volver() {
	}
/*	
	function ir(q){

		var q_seccion = ""
		if (q.indexOf("&seccion=")<0 ){
			q_seccion = "&seccion=<%=seccion%>"
		}
		var q_pag = ""
		if (q.indexOf("&pag=")<0 ){
			q_pag = "&pag=<%=pag%>"
		}
		location.href="index.asp?secc=<%=request.QueryString("secc")%>"+q+q_seccion+q_pag
	}


	function buscar() {
		ir("&cadena="+f.cadena.value)
	}
	function irPag(num) {
		ir("&pag="+num)
	}
	function ampliar(id) {
		ir("&ac=ampliar&id="+id)
	}
	function irSeccion(seccion, seccion2) {
		ir("&seccion="+seccion+"&seccion2="+seccion2+"&pag=1")
	}
	function todasSecciones() {
		ir("&seccion=&seccion2=")
	}
	function todasSecciones2() {
		ir("&seccion2=")
	}

	*/
	//-->
</script>


<%select case ac
case "ampliar" ' --------------------------------------------------------------------------- Ampliar noticia

	id = request.QueryString("id")&request.Form("id")
	if ""&id = "" or not isNumeric(id) then
		Response.Redirect(request.ServerVariables("HTTP_REFERER")&"&No se ha encontrado el registro solicitado.")
	end if
	


	if id <> "" then
		sql = ""
		sql = sql & "SELECT *"
		sql = sql & " FROM REGISTROS"
		sql = sql & " WHERE R_ID = " & id
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3
		re.Open()
		if re.eof then
			Response.Redirect(request.ServerVariables("HTTP_REFERER")&"&msg=No hay ningún registro con esta referencia.")
		else%>
		<table border="0" cellpadding="0" cellspacing="0" class="plantilla-tabla"  width="100%" align="center">
			<tr>
			  
			<td class="general-pixel-abajo">

			  <table width="100%"  border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td><%if nav_fecha and re("R_FECHA") <> 0 then%>
                    <table border="0" align="right" cellpadding="0" cellspacing="0">
                      <tr>
                        <td><span class="noticias-fecha"><%=re("R_FECHA")%></span></td>
                      </tr>
                    </table>
                    <%end if%>
                    <span class="noticias-titulo"><%=re("R_TITULO")%></span></td>
                </tr>
              </table>
			  </td>

			<td rowspan="4">&nbsp;</td>


			<%if nav_cuerpo <> "" then
				if re("R_"&nav_cuerpo) <> "" then%>

	      

	      <tr>
	        <td><span class="noticias-subtitulo"><br>
          <%=re("R_"&nav_cuerpo)%></span></td>
          </tr>
				<%end if
			end if%>
			
<%if zona = 2 then%>
		  <tr>
	        <td align="right" bgcolor="#3980F3">                <a target="_blank" href="/<%=c_s%>admin/aSkipper.asp?cualid=<%=cualid%>&direct=editar&id=<%=id%>&seccion=<%=re("R_SECCION")%>"><img src="/<%=c_s%>admin/global/img/lapiz.gif" alt=" Editar " border="0"></a><a target="_blank" href="/<%=c_s%>admin/aSkipper.asp?cualid=<%=cualid%>&direct=eliminar&id=<%=id%>"><img src="/<%=c_s%>admin/global/img/papelera.gif" alt=" Eliminar " border="0"></a>
           </td>
          </tr>
		  <%end if%>

		<tr>
		  <td>
		<br>
		<%if nav_foto and re("R_FOTO") <> "" then%>
			<table border='0' <%if ""&re("R_POS_FOTO") = "izq" then%>align="left"<%else%>align="right"<%end if%> cellpadding='0' cellspacing='0'>
				<tr>
				<td><img src='/<%=c_s%>img/foto_s_i.gif'></td>
				<td background='/<%=c_s%>img/foto_s.gif'><img src='/<%=c_s%>spacer.gif' width='1' height='1'></td>
				<td><img src='/<%=c_s%>img/foto_s_d.gif'></td>
				</tr>
				<tr>
				<td background='/<%=c_s%>img/foto_i.gif'><img src='/<%=c_s%>spacer.gif' width='1' height='1'></td>
				<td>
				<table cellpadding='0' cellspacing='0' border='0'>
					<tr>
					<td class='general-foto-pixel' align='center'><img src="/<%=c_s%>datos/<%=idioma%>/<%=cualid%>/fotos/<%=re("R_FOTO")%>"></td>
					</tr>
				</table>
				<%if re("R_PIE_FOTO") <> "" then%>
				<div align="center"><b><%=re("R_PIE_FOTO")%></b></div>
				<%end if%>
				
				</td>
				<td background='/<%=c_s%>img/foto_d.gif'><img src='/<%=c_s%>spacer.gif' width='1' height='1' border='0'></td>
				</tr>
				

				<tr>
				<td><img src='/<%=c_s%>img/foto_b_i.gif'></td>
				<td background='/<%=c_s%>img/foto_b.gif'><img src='/<%=c_s%>spacer.gif' width='1' height='1'></td>
				<td><img src='/<%=c_s%>img/foto_b_d.gif'></td>
				</tr>
			</table>
		<%end if%>

		<%
		' GENERACIÓN DE LOS CAMPOS CONFIGURADOS EN XML
		for each a in nodoCualid.childNodes
			c_nombre = ""&UCASE(a.getAttribute("campo"))
			c_titulo = ""&a.getAttribute("navtitulo")

			if a.nodeName = "dato" and c_nombre <> ucase(nav_cuerpo) then
			
				if re("R_"&c_nombre) <> "" then
					select case a.getAttribute("tipo")
				
					case "texto"

						if c_titulo <> "" then%>
							<b><%=c_titulo%></b>
							<br>
						<%end if%>
						<%=escribe( re("R_"&c_nombre) )%>
						<br><br>

					<%case "memo"

						if c_titulo <> "" then%>
							<b><%=c_titulo%></b>
							<br>
						<%end if%>
						<%=escribe( re("R_"&c_nombre) )%>
						<br><br>

					<%case "opcion"

						if c_titulo <> "" then%>
							<b><%=c_titulo%></b>
							<br>
						<%end if%>
						<%=escribe( re("R_"&c_nombre) )%>
						<br><br>
						
					<%case "combo"

						if c_titulo <> "" then%>
							<b><%=c_titulo%></b>
							<br>
						<%end if%>
						<%=escribe( re("R_"&c_nombre) )%>
						<br><br>

					<%case "check"

						if c_titulo <> "" then%>
							<b><%=c_titulo%></b>
							<br>
						<%end if%>
						<%=escribe( re("R_"&c_nombre) )%>
						<br><br>

					<%end select
				end if%>
					

			<%end if
		next%>


		<%if nav_fuente and re("R_FUENTE") <> "" then
			if nav_verenlace and re("R_ENLACE") <> "" then%>
				<br>Fuente: <a href="<%=escribe(re("R_ENLACE"))%>" target="_blank"><%=escribe(re("R_FUENTE"))%></a><br>
			<%else%>
				<br>Fuente: <%=escribe(re("R_FUENTE"))%><br>
			<%end if
		end if%>
		
		<%
		if nav_archivo then
			if re("R_ARCHIVO") <> "" then
			%>
			<br>
			<table border="0" cellpadding="2" cellspacing="0">
			  <tr>
			    <td><%pintaIconoExtension(re("R_TIPOARCHIVO"))%></td>
				<td><a href="/<%=c_s%>descargas/?idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>"><img src="/<%=c_s%>admin/global/img/descargar.gif" alt=" Descargar: <%=unpoco(re("R_TITULO"),60)%> " width="79" height="18" border="0" align="absmiddle"></a></td>
			  </tr>
			</table>

			<%
			end if
		end if
		%>
		
		</td>
		  </tr>
		</table>
		<br>
		<table width="100%" border="0" cellpadding="6" cellspacing="0">
		  <tr>
		    <td align="right">  <table border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td><table border="0" cellpadding="2" cellspacing="0" class="campo">
                  <tr>
                    <td><a href="<%=request.ServerVariables("HTTP_REFERER")%>" title="Volver a la p&aacute;gina anterior."> Atr&aacute;s</a> </td>
                  </tr>
                </table></td>
                <td><table border="0" cellpadding="2" cellspacing="0" class="campo">
                  <tr>
                    <td><a href="index.asp?secc=/<%=cualid%>" title="Ir al listado completo de registros.">Listado completo</a> </td>
                  </tr>
                </table></td>
              </tr>
            </table>
            </td>
          </tr>
	    </table>

  <%
		end if
		re.close
		set re = nothing
	end if

case else ' --------------------------------------------------------------------------- Listado, búsquedas ...%>


<%
' Cargar las secciones a un vector --
	sql = ""
	sql = sql & "SELECT *"
	sql = sql & " FROM SECCIONES"
	sql = sql & " WHERE S_REGISTROS > 0"
	sql = sql & " ORDER BY S_ORDEN "
	set reseccion = Server.CreateObject("ADODB.Recordset")
	on error resume next
		reseccion.ActiveConnection = conn_
		if err<>0 then
			unerror = true : msgerror = "No se ha podido cargar la base de datos para "&ucase(cualid)&" en el idioma "& ucase(idioma) &".<br>Conpruebe que la base de datos esté disponible en el idioma seleccionado."
		else
			reseccion.Source = sql : reseccion.CursorType = 3 : reseccion.CursorLocation = 2 : reseccion.LockType = 3
			reseccion.Open()
			if err<>0 then
				unerror = true : msgerror = "La consulta a la base de datos contiene algún error.<br>SQL:["&sql&"]"
			else
				Dim arrSeccionesId()
				Dim arrSeccionesNombre()
				Dim arrSeccionesTipo()
				if not 	reseccion.eof then
					numSecciones = 0
					while not reseccion.eof
						redim preserve arrSeccionesId(numSecciones) : arrSeccionesId(numSecciones) = reseccion("S_ID")
						redim preserve arrSeccionesNombre(numSecciones) : arrSeccionesNombre(numSecciones) = reseccion("S_NOMBRE")
						numSecciones = numSecciones + 1
						reseccion.movenext
					wend
				end if
			end if
		end if
		reseccion.Close()
		set reseccion = Nothing
	on error goto 0

	if not unerror then
	
		' Resetear paginado (pag=1) al hacer nueva búsqueda
		if StrComp(request.Form("cadenaanterior"),cadena) <>0 then
			pag = 1
		end if
		
		' CONSULTA ----------------------------------------------------------------------------------------------------
		sql = ""
		sql = sql & "SELECT *"
		sql = sql & " FROM REGISTROS, SECCIONES"
		sql = sql & " WHERE R_SECCION = S_ID"
		' Búsqueda
		if cadena <> "" then
			sql = sql & " AND (R_TITULO LIKE '%"& Replace(cadena,"'","''") & "%' OR R_MEMO1 LIKE '%"& Replace(cadena,"'","''") &"%')"
		end if
		
		if nav_activo then
			sql = sql & " AND R_ACTIVO = 1"		
		end if
	
		' Sección filtra
	
		if nav_filtraralias <> "" then
			if inStr(nav_filtraralias,",") then
			else
				sql = sql & " AND S_ALIAS = '"& nav_filtraralias &"'"
			end if
		end if
	
		if nav_activo_secciones then
			if numero(seccion) <> -1 then
				sql = sql & " AND R_SECCION = "& seccion
			end if
		end if
		if nav_activo_secciones2 then
			if numero(seccion2) <> -1 then
				sql = sql & " AND R_SECCION2 = "& seccion2
			end if
		end if
		
		if nav_orden <> "" then
			sql = sql & " ORDER BY "& nav_orden &""
		else
			sql = sql & " ORDER BY R_ORDEN ASC"
		end if
	
		'Response.Write("<br><fon size=1>"&sql&"</font>")
		' -------------------------------------------------------------------------------------------------------------
	
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3
		re.Open()
		if not re.eof then
			re.move ((pag * registrosPorPagina)) - registrosPorPagina
			totalRegistros = re.recordcount		
		else
			totalRegistros = 0
		end if

	end if ' unerror%>

<%if cadena <> "" then
	num = 0
	numPaginas = 0
	for n=1 to totalRegistros
		if n mod registrosPorPagina = 1 then
			num = num +1
			numPaginas = numPaginas + 1
		end if
	next%>
	<br>
		<%if totalRegistros > 0 then%>
	<table width="100%" border="0" cellpadding="4" cellspacing="0" bgcolor="f8f8f8">
		<tr>
		<td>
		

		<font color="#333333">Se han encontrado <b><%=totalRegistros%></b> resultados en la b&uacute;squeda para &quot;<font color="#5555aa"><b><%=cadena%></b></font>&quot;.</font>

		
		</td>
		<td align="right"> <font color="#333333">P&aacute;gina <%=pag%> de <%=numPaginas%></font>.</td>
		</tr>
	</table>
	<br>
		<%else%>
	<table width="100%" border="0" cellpadding="4" cellspacing="0" bgcolor="f8f8f8">
		<tr>
		<td>
		

<b>No se han encontrado resultados</b>		 </td>
		</tr>
	</table>
		
		<%end if%>
<%end if%>

<%if nav_activo_secciones and int(numSecciones) > 1 then%>

	<table width="100%" border="0" cellpadding="0" cellspacing="0">

		<tr>
		  <td align="right"><font color="#666666" size="1">Secciones</font></td>
		  </tr>
		<tr>
	<td class="noticias-secciones" align="left"><table border="0" cellpadding="2" cellspacing="3">
		<tr>
		<%for n=0 to numSecciones-1

				if int(arrSeccionesId(n)) = cint(seccion) then%>
					<td class="boton-over">&nbsp;<%=arrSeccionesNombre(n)%>&nbsp;</td>
				<%else%>
					<td class="boton-out"><a href="index.asp?secc=<%=secc%>&seccion=<%=arrSeccionesId(n)%>">&nbsp;<%=arrSeccionesNombre(n)%>&nbsp;</a></td>
				<%end if

		next
			if int(seccion) = -1 then%>
				<td class="boton-over">&nbsp;Todas&nbsp;</td>
			<%else%>
				<td class="boton-out">&nbsp;<a href="index.asp?secc=<%=secc%>&seccion=todas">Todas</a>&nbsp;</td>
			<%end if%>
		</tr>
		</table></td>
		</tr>
	</table>

<%end if

	sql = "SELECT * FROM SECCIONES2 WHERE S2_REGISTROS > 0 AND S2_ID_S = "& seccion &" ORDER BY S2_ORDEN "
	set re_subsecc = Server.CreateObject("ADODB.Recordset")
	re_subsecc.ActiveConnection = conn_
	re_subsecc.Source = sql : re_subsecc.CursorType = 3 : re_subsecc.CursorLocation = 2 : re_subsecc.LockType = 1
	re_subsecc.Open()
	
	numsubsecc=re_subsecc.recordcount
	
	if nav_activo_secciones2 and seccion <> -1 and numsubsecc>1 then%>
    <table  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
		

		<table border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td align="left">
                <table cellpadding="2" cellspacing="2">
                  <tr>
				  <td width="24" align="right" ><img src="../img/flecha_secc.gif" width="13" height="11"></td>
                    <%while not re_subsecc.eof
			if ""&seccion2 = ""&re_subsecc("S2_ID") then%>
                    <td width="17" class="boton-over"><b>&nbsp;<%=re_subsecc("S2_NOMBRE")%>&nbsp;</b></td>
                    <%else%>
                    <td width="25" class="boton-out"><a href="index.asp?secc=<%=secc%>&seccion=<%=seccion%>&seccion2=<%=re_subsecc("S2_ID")%>">&nbsp;<%=re_subsecc("S2_NOMBRE")%>&nbsp;</a></td>
                    <%end if%>
                    <%re_subsecc.movenext : wend
			if int(seccion) = -1 then%>
                    <td width="46" class="boton-over">&nbsp;Todas&nbsp;</td>
                    <%else%>
                    <td width="46" class="boton-out">&nbsp;<a href="index.asp?secc=<%=secc%>&seccion=todas&subseccion=todas">Todas</a>&nbsp;</td>
                    <%end if%>
                  </tr>
                </table>

            </td>
          </tr>
        </table></td>
      </tr>
    </table>
                <%
	re_subsecc.Close()
	set re_subsecc = Nothing
%>
    <br>
<%end if ' if seccion <> -1


if totalRegistros = 0 then%>
  <center>
  <b>No hay ning&uacute;n resultado disponible</b><br>
  <br>
  </center>
  <%else

	for n=0 to registrosPorPagina-1
		if not re.eof then
		
'			if re("R_ACTIVO") or not nav_activo then%>
		<table width="100%" border="0" cellpadding="4" cellspacing="0" class="noticias-fondo-lista">
		<tr>
		<%if nav_icono and re("R_ICONO") <> "" and (re("R_POS_ICONO") = "izq" or ""&re("R_POS_ICONO") = "") then%>
			<td width="1" valign="top"><img src="/<%=c_s%>datos/<%=idioma%>/<%=cualid%>/iconos/<%=re("R_ICONO")%>"><img src="../spacer.gif" width="1" height="1"></td>
		<%end if%>
		<td align="right" valign="top"><table width="100%" height="100%" border="0" cellpadding="2" cellspacing="0" bgcolor="ffffff">
		<tr> 
		<td align="left" class="general-pixel-abajo">
		<%
		if nav_archivo then
			if re("R_ARCHIVO") <> "" then
				pintaIconoExtension(re("R_TIPOARCHIVO"))
			end if
		end if
		%>
		
		<span class="titular-titulo">
		<%if nav_ampliar then%>
			<a href="../admin/visores/index.asp?secc=<%=secc%>&ac=ampliar&id=<%=re("R_ID")%>&seccion=<%=seccion%>&pag=<%=pag%>">
		<%end if%>
		
		<b><img src="../admin/img/flecha.gif" width="8" height="6" border="0">

		<%
			Response.Write re("R_TITULO")
			%></b>
		<%if nav_ampliar then%>
			</a>
		<%end if%>
		</span></td>
		<%if nav_fecha and re("R_FECHA") > 0 then%>
			<td class="general-pixel-abajo" align="right"><%=re("R_FECHA")%></td>
		<%end if%>
		</tr>
		<tr> 
		<td colspan="2" align="left" valign="top">
		<%if nav_cuerpo <> "" then
			if re("R_"&nav_cuerpo) <> "" then
				if nav_ampliar then
					Response.Write unpoco(re("R_"&nav_cuerpo),200)
				else
					Response.Write re("R_"&nav_cuerpo)
				end if
			else
				Response.Write "&nbsp;"
			end if
		end if
		
		if nav_archivo then
			if re("R_ARCHIVO") <> "" then
				%>
				<table border="0" align="right" cellpadding="2" cellspacing="0">
				<tr>
				<td><a href="../admin/descargas/?idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>"><img src="/<%=c_s%>admin/global/img/descargar.gif" alt=" Descargar: <%=unpoco(re("R_TITULO"),60)%> " width="79" height="18" border="0" align="absmiddle"></a></td>
				</tr>
				</table>
				<%
			end if
		end if%>


		</td>
		</tr>
		</table></td>
		<%if nav_icono and re("R_ICONO") <> "" and re("R_POS_ICONO") = "der" then%>
			<td width="1" valign="top"><img src="/<%=c_s%>datos/<%=idioma%>/<%=cualid%>/iconos/<%=re("R_ICONO")%>"><img src="../spacer.gif" width="1" height="1"></td>
		<%end if%>
		</tr>
		</table>
	 <br>
  <%'end if ' nav_activo
  re.movenext
			  
		end if
	next
end if
if totalRegistros >  registrosPorPagina then%>
	<div align="center">Páginas: 
	<%
	' Numero de página -------------------------------------------------------------------------
	Dim salida, num ,n
	salida = ""
	num = 0
	for n=1 to totalRegistros
		if n mod registrosPorPagina = 1 then
			num = num +1
			if int(num) = int(pag) then%>
				<b>[<%=num%>]</b>
			<%else%>
				<a href='../admin/visores/index.asp?secc=<%=secc%>&seccion=<%=seccion%>&seccion2=<%=seccion2%>&pag=<%=num%>'><%=num%></a>
			<%end if
		end if
	next
	' -------------------------------------------------------------------------------- Números de página%>
	</div>
<%else
	'Response.Write("<div align='center'>Paginas: <b>[1]</b></div>")
end if

re.Close() : set re = Nothing
end select ' ------------------------------------------------------------------------------------ end select%>
</form><!-- FIN DE FORMULARIO --------------------------------------------------------------------------------------- -->
<%if nav_busca then%>
	<script>f.cadena.focus()</script>
<%end if%>
</td>
</tr>
</table>
<%end if ' unerror


sub pintaIconoExtension (ext)
		select case ext
			case "jpg","png","gif","bmp","tif"
				ico = "img"
			case "exe"
				ico = "jpg"
			case "doc"
				ico = "jpg"
			case "txt"
				ico = "jpg"
			case "xls"
				ico = "jpg"
			case "zip"
				ico = "zip"
			case "pdf"
				ico = "pdf"
			case "mp3","wav"
				ico = "mp3"
			case else
				ico = "iconootro"
		end select
			%><img src="/<%=c_s%>img/<%=ico%>.gif" align="absmiddle" alt=" Archivo tipo: <%=UCASE(ext)%> "><%
	end sub

end if

if unerror then%>
	<font color="#FF0000"><b>ATENCIÓN</b>:<br>
<%=msgerror%></font>
<%end if%>