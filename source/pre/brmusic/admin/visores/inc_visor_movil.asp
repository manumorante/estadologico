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
<form name="f" action="index.asp?secc=<%=secc%>" method="POST" onSubmit="return envio()">
<input type="hidden" name="ac" value="<%=ac%>">
<input type="hidden" name="ref" value="">
<input type="hidden" name="id" value="">
<input type="hidden" name="pag" value="<%=pag%>">
<input type="hidden" name="seccion" value="<%=seccion%>">
<input type="hidden" name="seccion2" value="<%=seccion2%>">
<input type="hidden" name="seccion3" value="<%=seccion3%>">



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
		f.ac.value=""
		f.action="index.asp?secc=/<%=cualid%>"
		f.submit()
	}
-->
</script>

<%select case ac
case "ampliarfoto"%>
<img src="<%=request("foto")%>"><br>
<%
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
		
		<span class="noticias-titulo">
		<%if nav_fecha and re("R_FECHA") <> 0 then%>
			<div align="right"><%=re("R_FECHA")%></div>
		<%end if%>
		<%=re("R_TITULO")%>
		</span>
		
		<br>
		<%if nav_cuerpo <> "" then
			if re("R_"&nav_cuerpo) <> "" then%>
				<span><%=re("R_"&nav_cuerpo)%></span>
			<%end if%>
			<br>
			<%if nav_foto and re("R_FOTO") <> "" then
				if re("R_PIE_FOTO") <> "" then%>
					<div align="center"><b><%=re("R_PIE_FOTO")%></b></div>
				<%end if%>
				<a href="index.asp?secc=<%=secc%>&ac=ampliarfoto&foto=/<%=c_s%>datos/esp/<%=cualid%>/fotosmovil/<%=re("R_FOTO")%>">Ver foto</a>
				<br>
				<br>
			<%end if
			
			' GENERACI&Oacute;N DE LOS CAMPOS CONFIGURADOS EN XML
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
							<br>
						<%case "memo"
							if c_titulo <> "" then%>
								<b><%=c_titulo%></b>
								<br>
							<%end if%>
							<%=escribe( re("R_"&c_nombre) )%>
							<br>
						<%case "opcion"
							if c_titulo <> "" then%>
								<b><%=c_titulo%></b>
								<br>
							<%end if%>
							<%=escribe( re("R_"&c_nombre) )%>
							<br>
						<%case "combo"
							if c_titulo <> "" then%>
								<b><%=c_titulo%></b>
								<br>
							<%end if%>
							<%=escribe( re("R_"&c_nombre) )%>
							<br>
						<%case "check"
							if c_titulo <> "" then%>
								<b><%=c_titulo%></b>
								<br>
							<%end if%>
							<%=escribe( re("R_"&c_nombre) )%>
							<br>
						<%end select
					end if
				end if
			next
			if nav_fuente and re("R_FUENTE") <> "" then
				if nav_verenlace and re("R_ENLACE") <> "" then%>
					<br>
					Fuente: <a href="<%=escribe(re("R_ENLACE"))%>" target="_blank"><%=escribe(re("R_FUENTE"))%></a><br>
				<%else%>
					<br>
					Fuente: <%=escribe(re("R_FUENTE"))%><br>
				<%end if
			end if
			
			if nav_archivo then
				if re("R_ARCHIVO") <> "" then%>
					<br>
					<%pintaIconoExtension(re("R_TIPOARCHIVO"))%>
					<br>
					<a href="/<%=c_s%>descargas/?idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>"><img src="/<%=c_s%>admin/global/img/descargar.gif" alt=" Descargar: <%=unpoco(re("R_TITULO"),60)%> " width="79" height="18" border="0" align="absmiddle"></a>
				<%end if
			end if
			%><div align="right"><a href="index.asp?secc=/carrito&ac=add&idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>">A&ntilde;adir</a></div>
			<%
		end if
		re.close
		set re = nothing
	end if

end if

case else ' --------------------------------------------------------------------------- Listado, búsquedas ...

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
	next
	
	if totalRegistros > 0 then%>
		<br>
		Encontrado/s <b><%=totalRegistros%></b> resultados.
		<br>
		P&aacute;gina <%=pag%> de <%=numPaginas%>
		<br>
	<%else%>
		<br>
		<b>No se han encontrado resultados</b>
	<%end if
end if

if nav_activo_secciones and int(numSecciones) > 1 then

	for n=0 to numSecciones-1
		if int(arrSeccionesId(n)) = cint(seccion) then%>
			&nbsp;<span class="boton-out"><%=arrSeccionesNombre(n)%></span>&nbsp;
		<%else%>
			<a href="index.asp?secc=<%=secc%>&seccion=<%=arrSeccionesId(n)%>">&nbsp;<%=arrSeccionesNombre(n)%>&nbsp;</a>
		<%end if
	next
	
	if int(seccion) = -1 then%>
		&nbsp;<span class="boton-out">Todas</span>
	<%else%>
		&nbsp;<a href="index.asp?secc=<%=secc%>&seccion=todas">Todas</a>&nbsp;
	<%end if

end if

	sql = "SELECT * FROM SECCIONES2 WHERE S2_REGISTROS > 0 AND S2_ID_S = "& seccion &" ORDER BY S2_ORDEN "
	set re_subsecc = Server.CreateObject("ADODB.Recordset")
	re_subsecc.ActiveConnection = conn_
	re_subsecc.Source = sql : re_subsecc.CursorType = 3 : re_subsecc.CursorLocation = 2 : re_subsecc.LockType = 1
	re_subsecc.Open()
	
	numsubsecc=re_subsecc.recordcount
	
	if nav_activo_secciones2 and seccion <> -1 and numsubsecc>1 then
		while not re_subsecc.eof
			if ""&seccion2 = ""&re_subsecc("S2_ID") then%>
				<span class="boton-out"><%=re_subsecc("S2_NOMBRE")%></span>
			<%else%>
				<a href="index.asp?secc=<%=secc%>&seccion=<%=seccion%>&seccion2=<%=re_subsecc("S2_ID")%>"><%=re_subsecc("S2_NOMBRE")%></a>
			<%end if
			re_subsecc.movenext
		wend
		if int(seccion) = -1 then%>
			<span class="boton-out">Todas</span>
		<%else%>
			<a href="index.asp?secc=<%=secc%>&seccion=todas&subseccion=todas">Todas</a>
		<%end if
		re_subsecc.Close()
		set re_subsecc = Nothing%>
		<br>
	<%end if ' if seccion <> -1


if totalRegistros = 0 then%>
	<center>
	<b>No hay ning&uacute;n resultado </b><br>
	<br>
	</center>
<%else

	for n=0 to registrosPorPagina-1
		if not re.eof then%>
			<hr size="1" noshade>
			<%if nav_archivo then
				if re("R_ARCHIVO") <> "" then
					pintaIconoExtension(re("R_TIPOARCHIVO"))
				end if
			end if
			%>
			<%if nav_ampliar then%>
				<a href="index.asp?secc=<%=secc%>&ac=ampliar&id=<%=re("R_ID")%>&seccion=<%=seccion%>&pag=<%=pag%>">
			<%end if%>
			<b>&raquo;
			<%if nav_ampliar then
				Response.Write unpoco(re("R_TITULO"),60)
			else
				Response.Write re("R_TITULO")
			end if%>
			</b>
			<%if nav_ampliar then%>
				</a>
			<%end if%>
			<br>
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
				if re("R_ARCHIVO") <> "" then%>
					<div align="right"><a href="../descargas/?idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>">Descargar</a></div>
				<%end if
			end if%>
			<div align="right"><a href="index.asp?secc=/carrito&ac=add&idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>">A&ntilde;adir</a></div>
			<%re.movenext
		end if
	next%>
	<hr size="1" noshade>
	<%
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
				<a href='index.asp?secc=<%=secc%>&seccion=<%=seccion%>&seccion2=<%=seccion2%>&pag=<%=num%>'><%=num%></a>
			<%end if
		end if
	next
	' -------------------------------------------------------------------------------- Números de página%>
	</div>
<%end if

re.Close() : set re = Nothing
if nav_buscar then%>
<br>	
<div align="center"><br>
<table border="0" cellpadding="2" cellspacing="0" class="noticias-busqueda">
	<tr>
	<td valign="middle">
	<input name="cadenaanterior" type="hidden" value="<%=cadena%>">
	<input name="cadena" type="text" class="noticias-input" value="<%=cadena%>" size="15" maxlength="150">
	</td>
	<td valign="middle">
	<input type="submit" value="Buscar">
	</td>
	</tr>
</table>
	<a href="/<%=c_s%>esp/index.asp?secc=<%=secc%>"><br>
Eliminar b&uacute;squeda</a>
</div>

<%else%>
	<input type="hidden" name="cadena" value="">
	<input name="cadenaanterior" type="hidden" value="<%=cadena%>">
<%end if
end select ' ------------------------------------------------------------------------------------ end select%>
</form>
<%end if ' unerror




end if

if unerror then%>
	<font color="#FF0000"><b>ATENCIÓN</b>:<br>
<%=msgerror%></font>
<%end if%>