<% @LCID = 1034 %>
<html><!-- InstanceBegin template="/Templates/base2_eng.dwt.asp" codeOutsideHTMLIsLocked="false" --><head><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /><!-- InstanceBeginEditable name="doctitle" -->
<title>BR Music International .:. festivales de musica, nacionales e internacionales, escenarios, camerinos, vallas, catering, montaje fijación de carteleria, promoción y mailing, grupos electrógenos</title>
<!-- InstanceEndEditable --><link rel="stylesheet" type="text/css" href="arch/estilos.css"><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable --></head><body bgcolor="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"><div align="center"><table id="Tabla_01" width="776" height="580" border="0" cellpadding="0" cellspacing="0"><tr><td colspan="6"><img src="/brmusic/Templates/web_eng_01.jpg" width="775" height="22" alt=""></td><td><img src="espacio.gif" width="1" height="22" alt=""></td></tr><tr><td rowspan="16"><img src="/brmusic/Templates/web_eng_02.jpg" width="14" height="557" alt=""></td><td rowspan="2"><a href="/index_eng.asp"><img src="/brmusic/Templates/web_eng_03.jpg" alt="" width="126" height="111" border="0"></a></td>
<td colspan="4"><img src="/brmusic/Templates/web_eng_04.jpg" width="635" height="65" alt=""></td><td><img src="espacio.gif" width="1" height="65" alt=""></td></tr><tr><td rowspan="11"><img src="/brmusic/Templates/web_eng_05.jpg" width="37" height="326" alt=""></td><td colspan="2" rowspan="13" align="left" valign="top" background="/brmusic/Templates/web_eng_06.jpg" bgcolor="#EDE6C9"><!-- InstanceBeginEditable name="Cuerpo" -->
<%idioma = "esp" : cualid = "conciertos"%>
<!--#include file="visores/inc_conn.asp" -->
<%

	if not unerror then
		sql = "SELECT * FROM REGISTROS WHERE R_SECCION = 273 ORDER BY R_FECHA"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 1 : re.CursorLocation = 2 : re.LockType = 1 : re.Open()
	end if 'unerror
	
	if not unerror then

%>
	  <div style="overflow:auto; width:521; height:440;">

		<%if re.eof then%>
			<strong>No hay conciertos previstos</strong>
		<%else%>
			<table width="100%" border="0" cellspacing="0" cellpadding="10">
			<%while not re.eof%>
          <tr>
            <td><h3><%=re("R_TITULO")%></h3>

			<table border="0" cellspacing="0" cellpadding="2">

				<%if re("R_TEXT1") <> "" then%>
					<tr>
					  <td colspan="2" align="left" valign="middle"><strong><%=re("R_TEXT1")%></strong></td>
					</tr>
				<%end if%>

				<%if re("R_FECHA") <> "0:00:00" and re("R_FECHA") <> "" then%>
					<tr>
					  <td align="right" valign="middle"><span class="datos-concierto">DATE:</span></td>
					  <td valign="middle"><%=re("R_FECHA")%></td>
					</tr>
				<%end if%>

				<%if re("R_TEXT3") <> "" then%>
					<tr>
					  <td align="right" valign="middle"><span class="datos-concierto">VENUE:</span></td>
					  <td valign="middle"><%=re("R_TEXT3")%></td>
					</tr>
				<%end if%>

				<%if re("R_TEXT4") <> "" then%>
					<tr>
					  <td align="right" valign="middle"><span class="datos-concierto">CITY:</span></td>
					  <td valign="middle"><%=re("R_TEXT4")%></td>
					</tr>
				<%end if%>

			</table>
			
				<%if re("R_MEMO1") <> "" then%>
					<strong><font size="2"><%=re("R_MEMO1")%></font></strong><br>
				<%end if%>

				<%if re("R_TEXT5") <> "" then%>
					<span class="datos-concierto"><a href="<%="http://"& replace(re("R_TEXT5"),"http://","")%>" target="_blank">BUY YOUR TICKET</a></span>
				    <%end if%>
            <td align="right">
			<%if re("R_ICONO") <> "" then%>
			<table width="170" height="129" border="0" cellpadding="0" cellspacing="0" background="arch/bg_foto.jpg">
                <tr>
                  <td align="center" valign="middle"><img src="/datos/esp/conciertos/iconos/<%=re("R_ICONO")%>" alt="<%=re("R_TITULO")%>" width="156"></td>
                </tr>
            </table>
			<%end if%>
			</td>
          </tr>
          <tr>
            <td colspan="2"><table width="100%" height="3" border="0" cellpadding="0" cellspacing="0" bgcolor="#E4D6AF">
                <tr>
                  <td><img src="arch/espacio.gif" width="100%" height="1"></td>
                </tr>
            </table></td>
          </tr>
		  <%re.movenext
		  wend%>
        </table>
		<%end if%>

      </div>
	<%end if ' unerror

	on error resume next
	re.Close()
	set re = nothing
	on error goto 0
	%>
	<!--#include file="inc_alerta.asp" -->

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
<td><img src="espacio.gif" width="1" height="62" alt=""></td></tr><tr><td colspan="2"><img src="/brmusic/Templates/web_eng_20.jpg" width="521" height="29" alt=""></td><td><img src="espacio.gif" width="1" height="29" alt=""></td></tr><tr><td><img src="/brmusic/Templates/web_eng_21.jpg" width="489" height="23" alt=""></td><td colspan="2"><a href="http://www.estadologico.com/"><img src="/brmusic/Templates/web_eng_22.jpg" alt="Diseño Web Granada - Estado Lógico" width="109" height="23" border="0"></a></td>
<td><img src="espacio.gif" width="1" height="23" alt=""></td></tr><tr><td><img src="espacio.gif" width="14" height="1" alt=""></td><td><img src="espacio.gif" width="126" height="1" alt=""></td><td><img src="espacio.gif" width="37" height="1" alt=""></td><td><img src="espacio.gif" width="489" height="1" alt=""></td><td><img src="espacio.gif" width="32" height="1" alt=""></td><td><img src="espacio.gif" width="77" height="1" alt=""></td><td></td></tr></table></div></body><!-- InstanceEnd --></html>
