<%


	Dim unerror, msgerror
	conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("\"& c_s &"datos\"& idioma &"\"& cualid &"\"& cualid &".mdb")
	on error resume next
		set conn_activa = server.CreateObject("ADODB.Connection")
		conn_activa.Open conn_
		if err<>0 then
			unerror = true : msgerror = "No hay conexión."
		end if
	on error goto 0

%>
<!--#include virtual="/admin/global/inc_inicia_xml.asp" -->
<!--#include virtual="/admin/global/inc_rutinas.asp" -->
<%

	' Paginado
	Dim registrosPorPagina
	registrosPorPagina = 5

	'Inicio (Cabeza lectora)
	Dim pag
	pag = numero(request.QueryString("pag"))
	if pag = 0 then pag = 1 end if

	dim ac ' Acción
	dim id	' Identificador general
	dim reTotal, re ' (consultaX)
	dim idsecc	' Id de la sección
	dim idsecc2	' Id de la sub sección
	
	ac = ""&request.QueryString("ac")
	id = numero(request.QueryString("id"))
	idsecc = numero(request.QueryString("idsecc"))
	idsecc2 = numero(request.QueryString("idsecc2"))
	
	select case ac
	case "info"
		sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
		consultaXOpen sql,1
			if reTotal > 0 then%>
				<div align="left">
				<h1><%=re("R_TITULO")%></h1>
				<hr size="1" noshade/>
				<%=re("R_MEMO1")%>
				<%if ""& re("R_FOTO") <> "" then%>
				<p><a href="/<%=c_s%>datos/<%=idioma%>/<%=cualid%>/fotosmovil/<%=re("R_FOTO")%>">Ver foto ampliada</a></p>
				<%end if%>
				<hr size="1" noshade/>
				</div>
				<a href="index.asp?secc=/carrito&ac=opt&idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>&seccrefer=<%=secc%>&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>&pag=<%=request.QueryString("pag")%>">A&ntilde;adir</a><br />
			<%end if
		consultaXCLose()
		
	case "secc"

		sql = "SELECT * FROM SECCIONES"
		consultaXOpen sql,1
			if reTotal >0 then
				while not re.eof
					if re("S_REGISTROS") >0 then
						if re("S_SUBSECCIONES") >0 then%>
							<a href="index.asp?secc=<%=secc%>&ac=secc2&id=<%=re("S_ID")%>"><%=re("S_NOMBRE")%></a><br />
						<%else%>
							<a href="index.asp?secc=<%=secc%>&ac=listado&idsecc=<%=re("S_ID")%>"><%=re("S_NOMBRE")%></a><br />
						<%end if
					end if%>
					<%re.movenext
				wend
			end if
		consultaXCLose()
		%>
		<br />
		<div align="left">
		<form name="buscar" action="index.asp" method="get">
		<input type="hidden" name="secc" value="<%=secc%>" />
		<input type="hidden" name="ac" value="listado" />
		<input type="hidden" name="idsecc" value="<%=request.QueryString("idsecc")%>" />
		<input type="hidden" name="idsecc2" value="<%=request.QueryString("idsecc2")%>" />
		<input type="text" name="cadena" value="<%=request.QueryString("cadena")%>"/>
		<input name="" type="submit" value="Buscar" />
		</form>
		<a href="index.asp?secc=<%=secc%>&ac=listado">Listado completo</a><br />
		<a href="index.asp?secc=<%=secc%>&pag=<%=request.QueryString("pag")%>">Inicio</a>
		</div>
		<%case "secc2"

		consultaXOpen "SELECT S_NOMBRE FROM SECCIONES WHERE S_ID="& id,1
			if reTotal >0 then%>
				<B><%=re("S_NOMBRE")%></B>
				<br />
			<%end if
			reTotal = 0
		consultaXCLose()

		consultaXOpen "SELECT * FROM SECCIONES2 WHERE S2_ID_S="& id,1
			if reTotal >0 then
				while not re.eof
					if re("S2_REGISTROS")>0 then%>
						&raquo;<a href="index.asp?secc=<%=secc%>&ac=listado&idsecc=<%=id%>&idsecc2=<%=re("S2_ID")%>"><%=re("S2_NOMBRE")%></a>
						<br />
					<%end if
					re.movenext
				wend
			end if
		consultaXCLose()
		%>
		<form name="buscar" action="index.asp" method="get">
          <input type="hidden" name="secc" value="<%=secc%>" />
          <input type="hidden" name="ac" value="listado" />
          <input type="hidden" name="idsecc" value="<%=request.QueryString("id")%>" />
          <input type="text" name="cadena" value="<%=request.QueryString("cadena")%>"/>
          <input name="" type="submit" value="Buscar" />
                </form>
		<%

	case "listado"
	
		' Listado
		sql = "SELECT * FROM REGISTROS"
		sql = sql & " WHERE R_ID>0"

		' Sección
		if idsecc >0 then
			sql = sql & " AND R_SECCION="& idsecc
		end if
		' Sección 2
		if idsecc2 >0 then
			sql = sql & " AND R_SECCION2="& idsecc2
		end if
		
		' Búsqueda
		cadena = ""& request.QueryString("cadena")
		if cadena <> "" then
			sql = sql & " AND (R_TITULO LIKE '%"& replace(cadena,"'","''") &"%' OR R_REF LIKE '%"& replace(cadena,"'","''") &"%')"
		end if

		consultaXOpen sql,1
			if reTotal >0 then
				re.move (pag * registrosPorPagina) - registrosPorPagina
				for n=0 to registrosPorPagina-1
					if not re.eof then%>
						<div align="left">&raquo; <a href="index.asp?secc=<%=secc%>&ac=info&id=<%=re("R_ID")%>&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>&pag=<%=request.QueryString("pag")%>"><%=re("R_TITULO")%></a> <%=re("R_PRECIO")%> &euro;
						<br />
						<%=re("R_TEXT1")%><br />
						</div>
						<div align="right"><a href="index.asp?secc=/carrito&ac=opt&idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>&seccrefer=<%=secc%>&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>&pag=<%=request.QueryString("pag")%>">A&ntilde;adir</a></div>
						<hr size="1" noshade/>
						<%re.movenext
					end if
				next%>
			<%end if
		consultaXClose()

		if reTotal >  registrosPorPagina then%>
			<div align="center">Páginas: 
			<%
			' Numero de página -------------------------------------------------------------------------
			salida = ""
			num = 0
			for n=1 to reTotal
				if n mod registrosPorPagina = 1 then
					num = num +1
					if int(num) = int(pag) then%>
						<b>[<%=num%>]</b>
					<%else%>
						<a href='index.asp?secc=<%=secc%>&ac=listado&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>&pag=<%=num%>'><%=num%></a>
					<%end if
				end if
			next
			' -------------------------------------------------------------------------------- Números de página%>
			</div>
		<%end if%>
		
<div align="left">
		<form name="buscar" action="index.asp" method="get">
		<input type="hidden" name="secc" value="<%=secc%>" />
		<input type="hidden" name="ac" value="listado" />
		<input type="hidden" name="idsecc" value="<%=request.QueryString("idsecc")%>" />
		<input type="hidden" name="idsecc2" value="<%=request.QueryString("idsecc2")%>" />
		<input type="text" name="cadena" value="<%=request.QueryString("cadena")%>"/>
		<input name="" type="submit" value="Buscar" />
		<a href="index.asp?secc=<%=secc%>&ac=listado&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>">Eliminar b&uacute;squeda</a>
		</form>
		
		<a href="index.asp?secc=/carrito&seccrefer=<%=secc%>&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>&pag=<%=request.QueryString("pag")%>">Ver carrito</a><br />
		<a href="index.asp?secc=<%=secc%>&ac=secc&pag=<%=request.QueryString("pag")%>">Secciones</a><br />
		<a href="index.asp?secc=/iniciomovil">Men&uacute; principal</a><br />
  <a href="index.asp?secc=<%=secc%>&ac=listado&pag=<%=request.QueryString("pag")%>">Listado completo</a>
</div>
	<%case else%>
		<div align="left">
		<form name="buscar" action="index.asp" method="get">
		<input type="hidden" name="secc" value="<%=secc%>" />
		<input type="hidden" name="ac" value="listado" />
		<input type="hidden" name="idsecc" value="<%=request.QueryString("idsecc")%>" />
		<input type="hidden" name="idsecc2" value="<%=request.QueryString("idsecc2")%>" />
		<input type="text" name="cadena" value="<%=request.QueryString("cadena")%>"/>
		<input name="" type="submit" value="Buscar" />
		</form><a href="index.asp?secc=<%=secc%>&ac=secc&pag=<%=request.QueryString("pag")%>">Secciones</a><br />
		<a href="index.asp?secc=<%=secc%>&ac=listado&pag=<%=request.QueryString("pag")%>">Listado completo</a><br />
		<a href="index.asp?secc=/iniciomovil">Men&uacute; principal</a><br />
		</div>
	<%end select%>