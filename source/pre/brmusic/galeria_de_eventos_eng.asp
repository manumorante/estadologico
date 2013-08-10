<% @LCID = 1034 %>
<html><!-- InstanceBegin template="/Templates/base2_eng.dwt.asp" codeOutsideHTMLIsLocked="false" --><head><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /><!-- InstanceBeginEditable name="doctitle" -->
<title>BR Music International .:. festivales de musica, nacionales e internacionales, escenarios, camerinos, vallas, catering, montaje fijaci&oacute;n de carteleria, promoci&oacute;n y mailing, grupos electr&oacute;genos</title>
<!-- InstanceEndEditable --><link rel="stylesheet" type="text/css" href="arch/estilos.css"><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable --></head><body bgcolor="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"><div align="center"><table id="Tabla_01" width="776" height="580" border="0" cellpadding="0" cellspacing="0"><tr><td colspan="6"><img src="/brmusic/Templates/web_eng_01.jpg" width="775" height="22" alt=""></td><td><img src="espacio.gif" width="1" height="22" alt=""></td></tr><tr><td rowspan="16"><img src="/brmusic/Templates/web_eng_02.jpg" width="14" height="557" alt=""></td><td rowspan="2"><a href="/index_eng.asp"><img src="/brmusic/Templates/web_eng_03.jpg" alt="" width="126" height="111" border="0"></a></td>
<td colspan="4"><img src="/brmusic/Templates/web_eng_04.jpg" width="635" height="65" alt=""></td><td><img src="espacio.gif" width="1" height="65" alt=""></td></tr><tr><td rowspan="11"><img src="/brmusic/Templates/web_eng_05.jpg" width="37" height="326" alt=""></td><td colspan="2" rowspan="13" align="left" valign="top" background="/brmusic/Templates/web_eng_06.jpg" bgcolor="#EDE6C9"><!-- InstanceBeginEditable name="Cuerpo" -->
<div id="galeria-eventos" style="overflow:auto; width:521; height:440;">
<%idioma = "esp" : cualid = "conciertos"%>
<!--#include file="visores/inc_conn.asp" -->
<%
	' Listado de fechas:
	if not unerror then

		anteriores = false
		if ""& request.QueryString("anteriores") = "True" then
			anteriores = true
		end if

		if anteriores then
			anteriores_sql = anteriores_sql & " AND R_FECHA < #01/01/1998#"
		else
			anteriores_sql = anteriores_sql & " AND R_FECHA > #31/12/1997#"
		end if

		sql = "SELECT * FROM REGISTROS WHERE R_SECCION = 276 AND (R_TEXT6 = 'concierto')"
		sql = sql & anteriores_sql
		sql = sql & " ORDER BY R_FECHA DESC"
		set re = Server.CreateObject("ADODB.Recordset") : re.ActiveConnection = conn_ : re.Source = sql : re.CursorType = 1 : re.CursorLocation = 2 : re.LockType = 1 : re.Open()

		' Recuperar sólo los años, y sólo uno de cada.
		str_anos = ","
		while not re.eof
			fecha = CDate(re("R_FECHA"))
			ano = year(fecha)

			' Si el ano no está ya en la lista 
			if inStr(str_anos,","& ano &",") = 0 then
				if anteriores then
					' Anteriores
					if ano <=1997 then
						str_anos = str_anos & ano &","
					end if
				else
					' Posteriores
					if ano >1997 then
						str_anos = str_anos & ano &","
					end if
				end if

			end if
			re.movenext
		wend
		
		str_anos = left(str_anos,len(str_anos)-1)

		re.close() : set re = nothing		
	end if

	' Año elegido
	sql_ano_select = ""
	ano_select = ""& request.QueryString("year")
	if ano_select <> "" then
		ano_select = Cdate("1/1/"& ano_select)
		ano_mas = ano_select+365
		ano_menos = ano_select-1

		sql_ano_select = " AND (R_FECHA < #"& ano_mas &"# AND R_FECHA > #"& ano_menos &"#)"
	end if

	' Listado de conciertos
	if not unerror then
		sql = "SELECT * FROM REGISTROS WHERE R_SECCION = 276 AND (R_TEXT6 = 'concierto')"
		sql = sql & sql_ano_select
		sql = sql & anteriores_sql
		sql = sql &" ORDER BY R_FECHA DESC"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 1 : re.CursorLocation = 2 : re.LockType = 1 : re.Open()
	end if 'unerror
	
	if not unerror then

%>

		<p><span class="titulo-seccion">Events galery</span><br>
		Go to  <a href="#festivales">festivales</a>.		</p>

		<%if re.eof then%>
			<strong>No hay conciertos introducidos aún</strong>
		<%else

			registros = re.recordCount

			arr_anos = split(str_anos,",")

			%>

			<div id="anos">Year:<br> 
			  <%for each ano in arr_anos%>
					<a href="galeria_de_eventos.asp?year=<%=ano%>&anteriores=<%=anteriores%>"><%=ano%></a>				
				<%next%>
				<%if not anteriores then%>
		        <a href="galeria_de_eventos.asp?anteriores=True"><strong>Previous</strong></a>
				<%else%>
		        <a href="galeria_de_eventos.asp"><strong>Following</strong></a>
				<%end if%>
		        <a href="#festivales"><strong>Festivales</strong></a>
		  </div>

			<%while not re.eof

				fecha = Cdate(re("R_FECHA"))
				ano = year(fecha)
				
				if ano <> temp then%>
					<div class="corte-ano"><a name="year<%=ano%>"></a><%=ano%></div>
					<%
					temp = ano
				end if%>

				<div class="evento">
					<img src="/datos/esp/conciertos/fotos/<%=re("R_FOTO")%>" alt="<%=re("R_TITULO")%>" border="0" />
					<div class="datos">
						<h2 class="titulo-evento"><%=re("R_TITULO")%></h2>
						<%if re("R_TEXT1") <> "" then%>
							<span class="tit-grupos">Grupos:</span><br>
							<span class="grupos"><%=re("R_TEXT1")%></span><br>
						<%end if%>
						<span class="tit-recinto">Venue:</span><br>
						<span class="recinto"><%=re("R_TEXT3")%></span><br>
						<span class="tit-lugar">City:</span><br>
						<span class="lugar"><%=re("R_TEXT4")%></span>
					</div>
				</div>
				<%
				re.movenext
			wend

		end if%>

	<%end if ' unerror

	on error resume next
	re.Close()
	set re = nothing
	on error goto 0

	
	

	' Listado de festivales
	if not unerror then
		sql = "SELECT * FROM REGISTROS WHERE R_SECCION = 276 AND (R_TEXT6 = 'festival')"& sql_ano_select &" ORDER BY R_FECHA DESC"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 1 : re.CursorLocation = 2 : re.LockType = 1 : re.Open()
	end if 'unerror
	
	if not unerror then

%>

		<p><span class="titulo-seccion"><a name="festivales"></a>Festivales</span></p>

		<%if re.eof then%>
			<strong>No hay festivales introducidos aún</strong>
		    <%else

			registros = re.recordCount

			arr_anos = split(str_anos,",")

		dim cuenta
		cuenta = 0
		while not re.eof

				fecha = Cdate(re("R_FECHA"))
				ano = year(fecha)
				
				if ano <> temp then%>
		<%
					temp = ano
				end if%>

				<div class="evento">
					<div class="foto">
					<%
					if re("R_FOTO") <> "" then
						%>
					<img src="/datos/esp/conciertos/fotos/<%=re("R_FOTO")%>" alt="<%=re("R_TITULO")%>" border="0" />
					<%
					end if
					%>
					</div>
					<div class="datos">
						<h2 class="titulo-evento"><%=re("R_TITULO")%></h2>
						<%if re("R_TEXT3") <> "" then
							%>
						<span class="tit-recinto">Venue:</span><br>
							<span class="recinto"><%=re("R_TEXT3")%></span><br><%
						end if

						if re("R_TEXT4") <> "" then
							%>
							<span class="tit-lugar">City:</span><br>
							<span class="lugar"><%=re("R_TEXT4")%></span><br><%
						end if

						if re("R_TEXT1") <> "" then
							%>
							<a href="galeria_de_eventos_ficha_eng.asp?id=<%=re("R_ID")%>">&raquo; Attending groups</a>
							<%
						end if
						%></div>
		  </div><%
				re.movenext
				cuenta = cuenta + 1
				if not cbool(cuenta mod 3) then
					%><div class="corte-festival">&nbsp;</div><%
				end if

			wend

		end if%>

	<%end if ' unerror

	on error resume next
	re.Close()
	set re = nothing
	on error goto 0
	%>

<!--#include file="inc_alerta.asp" -->

</div>

    <!-- InstanceEndEditable --></td><td rowspan="14"><img src="/brmusic/Templates/web_eng_07.jpg" width="77" height="469" alt=""></td><td><img src="espacio.gif" width="1" height="46" alt=""></td></tr><tr><td><img src="/brmusic/Templates/web_eng_08.jpg" width="126" height="24" alt=""></td><td><img src="espacio.gif" width="1" height="24" alt=""></td></tr><tr><td><%
	  url = request.ServerVariables("URL")
	  %><a href="/index_eng.asp"><%
		if inStr(url,"index_eng.asp") then
	  		%><img src="/brmusic/Templates/web2_eng_09.jpg" alt="" width="126" height="20" border="0"><%
		else
			%><img src="/brmusic/Templates/web_eng_09.jpg" alt="" width="126" height="20" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="20" alt=""></td></tr><tr><td><a href="galeria_de_eventos_eng.asp"><%
		if inStr(url,"/galeria_de_eventos_eng.asp") then
			%><img src="/brmusic/Templates/web2_eng_10.jpg" alt="" width="126" height="24" border="0"><%
		else
			%><img src="/brmusic/Templates/web_eng_10.jpg" alt="" width="126" height="24" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="24" alt=""></td></tr><tr><td><a href="servicios_eng.asp"><%
		if inStr(url,"/servicios_eng.asp") then
			%><img src="/brmusic/Templates/web2_eng_11.jpg" alt="" width="126" height="20" border="0"><%
		else
			%><img src="/brmusic/Templates/web_eng_11.jpg" alt="" width="126" height="20" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="20" alt=""></td></tr><tr><td><a href="contratacion_eng.asp"><%
		if inStr(url,"/contratacion_eng.asp") then
			%><img src="/brmusic/Templates/web2_eng_12.jpg" alt="" width="126" height="21" border="0"><%
		else
			%><img src="/brmusic/Templates/web_eng_12.jpg" alt="" width="126" height="21" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="21" alt=""></td></tr><tr><td><a href="prensa_eng.asp"><%
		if inStr(url,"/prensa_eng.asp") then
			%><img src="/brmusic/Templates/web2_eng_13.jpg" alt="" width="126" height="21" border="0"><%
		else
			%><img src="/brmusic/Templates/web_eng_13.jpg" alt="" width="126" height="21" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="21" alt=""></td></tr><tr><td><a href="contacto_eng.asp"><%
 		if inStr(url,"/contacto_eng.asp") then
			%><img src="/brmusic/Templates/web2_eng_14.jpg" alt="" width="126" height="22" border="0"><%
		else
			%><img src="/brmusic/Templates/web_eng_14.jpg" alt="" width="126" height="22" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="22" alt=""></td></tr><tr><td><a href="trabaja_con_nosotros_eng.asp"><%
		if inStr(url,"/trabaja_con_nosotros_eng.asp") then
			%><img src="/brmusic/Templates/web2_eng_15.jpg" alt="" width="126" height="22" border="0"><%
		else
			%><img src="/brmusic/Templates/web_eng_15.jpg" alt="" width="126" height="22" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="22" alt=""></td></tr><tr><td><a href="lista_de_correo_eng.asp"><%
		if inStr(url,"/lista_de_correo_eng.asp") then
			%><img src="/brmusic/Templates/web2_eng_16.jpg" alt="" width="126" height="25" border="0"><%
		else
			%><img src="/brmusic/Templates/web_eng_16.jpg" alt="" width="126" height="25" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="25" alt=""></td></tr><tr><td><img src="/brmusic/Templates/web_eng_17.jpg" width="126" height="81" alt=""></td><td><img src="espacio.gif" width="1" height="81" alt=""></td></tr><tr><td colspan="2"><a href="http://www.brmusic.net/Festival_Atarfe_Vega_Rock/" target="_blank"><img src="/brmusic/Templates/web_eng_18.jpg" alt="Atarfe Vega Rock" width="163" height="52" border="0"></a></td><td><img src="espacio.gif" width="1" height="52" alt=""></td></tr><tr><td colspan="2" rowspan="3"><a href="http://www.onroadtour.com.br/" target="_blank"><img src="/brmusic/Templates/web_eng_19.jpg" alt="On Road Tour" width="163" height="114" border="0"></a></td>
<td><img src="espacio.gif" width="1" height="62" alt=""></td></tr><tr><td colspan="2"><img src="/brmusic/Templates/web_eng_20.jpg" width="521" height="29" alt=""></td><td><img src="espacio.gif" width="1" height="29" alt=""></td></tr><tr><td><img src="/brmusic/Templates/web_eng_21.jpg" width="489" height="23" alt=""></td><td colspan="2"><a href="http://www.estadologico.com/"><img src="/brmusic/Templates/web_eng_22.jpg" alt="Dise&ntilde;o Web Granada - Estado L&oacute;gico" width="109" height="23" border="0"></a></td>
<td><img src="espacio.gif" width="1" height="23" alt=""></td></tr><tr><td><img src="espacio.gif" width="14" height="1" alt=""></td><td><img src="espacio.gif" width="126" height="1" alt=""></td><td><img src="espacio.gif" width="37" height="1" alt=""></td><td><img src="espacio.gif" width="489" height="1" alt=""></td><td><img src="espacio.gif" width="32" height="1" alt=""></td><td><img src="espacio.gif" width="77" height="1" alt=""></td><td></td></tr></table></div></body><!-- InstanceEnd --></html>
