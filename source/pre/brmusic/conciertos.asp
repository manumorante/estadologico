<html><!-- InstanceBegin template="/Templates/base2.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!--#include file="datos/inc_config_gen.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>BR Music International .:. festivales de musica, nacionales e internacionales, escenarios, camerinos, vallas, catering, montaje fijación de carteleria, promoción y mailing, grupos electrógenos</title>
<!-- InstanceEndEditable -->
<link rel="stylesheet" type="text/css" href="arch/estilos.css">
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
</head>
<body bgcolor="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="center"><table id="Tabla_01" width="776" height="580" border="0" cellpadding="0" cellspacing="0"><tr><td colspan="7"><img src="/brmusic/Templates/web_01.jpg" width="775" height="22" alt=""></td><td><img src="espacio.gif" width="1" height="22" alt=""></td></tr><tr><td rowspan="18"><img src="/brmusic/Templates/web_02.jpg" width="14" height="557" alt=""></td><td colspan="2" rowspan="2"><a href="/brmusic/"><img src="/brmusic/Templates/web_03.jpg" width="126" height="111" border="0"></a></td>
<td colspan="4"><img src="/brmusic/Templates/web_04.jpg" width="635" height="65" alt=""></td><td><img src="espacio.gif" width="1" height="65" alt=""></td></tr><tr><td rowspan="13"><img src="/brmusic/Templates/web_05.jpg" width="37" height="326" alt=""></td><td colspan="2" rowspan="15" align="left" valign="top" background="/brmusic/Templates/web_06.jpg" bgcolor="#EDE6C9"><!-- InstanceBeginEditable name="Cuerpo" -->
<%idioma = "esp" : cualid = "conciertos"%>
<!--#include file="visores/inc_conn.asp" -->
<%

	if not unerror then
		sql = "SELECT * FROM REGISTROS ORDER BY R_ORDEN_SECCION"
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
                <strong><%=re("R_TEXT1")%></strong><br>
                <span class="datos">DíA:</span> <%=re("R_TEXT2")%><br>
                <span class="datos">SALA:</span> <%=re("R_TEXT3")%><br>
                <span class="datos">LUGAR:</span> <%=re("R_TEXT4")%><br></td>
            <td align="right">
			<%if re("R_FOTO") <> "" then%>
			<table width="170" height="129" border="0" cellpadding="0" cellspacing="0" background="arch/bg_foto.jpg">
                <tr>
                  <td align="center" valign="middle"><img src="/<%=c_s%>datos/esp/conciertos/foto/<%=re("R_FOTO")%>" alt="<%=re("R_TITULO")%>" width="156"></td>
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

    <!-- InstanceEndEditable --></td><td rowspan="16"><img src="/brmusic/Templates/web_07.jpg" width="77" height="469" alt=""></td><td><img src="espacio.gif" width="1" height="46" alt=""></td></tr><tr><td colspan="2"><img src="/brmusic/Templates/web_08.jpg" width="126" height="24" alt=""></td><td><img src="espacio.gif" width="1" height="24" alt=""></td></tr><tr><td colspan="2"><%
	  url = request.ServerVariables("URL")
	  %><a href="/brmusic/"><%
		if inStr(url,"/brmusic/index.asp") then
	  		%><img src="/brmusic/Templates/web2_09.jpg" alt="" width="126" height="20" border="0"><%
		else
			%><img src="/brmusic/Templates/web_09.jpg" alt="" width="126" height="20" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="20" alt=""></td></tr><tr><td colspan="2"><a href="galeria_de_eventos.asp"><%
		if inStr(url,"/galeria_de_eventos.asp") then
			%><img src="/brmusic/Templates/web2_10.jpg" alt="" width="126" height="34" border="0"><%
		else
			%><img src="/brmusic/Templates/web_10.jpg" alt="" width="126" height="34" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="34" alt=""></td></tr><tr><td colspan="2"><a href="servicios.asp"><%
		if inStr(url,"/servicios.asp") then
			%><img src="/brmusic/Templates/web2_11.jpg" alt="" width="126" height="21" border="0"><%
		else
			%><img src="/brmusic/Templates/web_11.jpg" alt="" width="126" height="21" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="21" alt=""></td></tr><tr><td colspan="2"><a href="contratacion.asp"><%
		if inStr(url,"/contratacion.asp") then
			%><img src="/brmusic/Templates/web2_12.jpg" alt="" width="126" height="22" border="0"><%
		else
			%><img src="/brmusic/Templates/web_12.jpg" alt="" width="126" height="22" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="22" alt=""></td></tr><tr><td colspan="2"><a href="prensa.asp"><%
		if inStr(url,"/prensa.asp") then
			%><img src="/brmusic/Templates/web2_13.jpg" alt="" width="126" height="21" border="0"><%
		else
			%><img src="/brmusic/Templates/web_13.jpg" alt="" width="126" height="21" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="21" alt=""></td></tr><tr><td colspan="2"><a href="contacto.asp"><%
 		if inStr(url,"/contacto.asp") then
			%><img src="/brmusic/Templates/web2_14.jpg" alt="" width="126" height="21" border="0"><%
		else
			%><img src="/brmusic/Templates/web_14.jpg" alt="" width="126" height="21" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="21" alt=""></td></tr><tr><td colspan="2"><a href="trabaja_con_nosotros.asp"><%
		if inStr(url,"/trabaja_con_nosotros.asp") then
			%><img src="/brmusic/Templates/web2_15.jpg" alt="" width="126" height="31" border="0"><%
		else
			%><img src="/brmusic/Templates/web_15.jpg" alt="" width="126" height="31" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="31" alt=""></td></tr><tr><td colspan="2"><a href="lista_de_correo.asp"><%
		if inStr(url,"/lista_de_correo.asp") then
			%><img src="/brmusic/Templates/web2_16.jpg" alt="" width="126" height="31" border="0"><%
		else
			%><img src="/brmusic/Templates/web_16.jpg" alt="" width="126" height="31" border="0"><%
		end if
		%></a></td><td><img src="espacio.gif" width="1" height="31" alt=""></td></tr><tr><td colspan="2"><img src="/brmusic/Templates/web_17.jpg" width="126" height="16" alt=""></td><td><img src="espacio.gif" width="1" height="16" alt=""></td></tr><tr><td rowspan="2"><img src="/brmusic/Templates/web_18.jpg" width="56" height="39" alt=""></td><td><a href="/index_eng.asp"><img src="/brmusic/Templates/web_19.jpg" alt="English" width="70" height="23" border="0"></a></td>
<td><img src="espacio.gif" width="1" height="23" alt=""></td></tr><tr><td><img src="/brmusic/Templates/web_20.jpg" width="70" height="16" alt=""></td><td><img src="espacio.gif" width="1" height="16" alt=""></td></tr><tr><td colspan="3"><a href="http://www.brmusic.net/Festival_Atarfe_Vega_Rock/" target="_blank"><img src="/brmusic/Templates/web_21.jpg" alt="" width="163" height="52" border="0"></a></td><td><img src="espacio.gif" width="1" height="52" alt=""></td></tr><tr><td colspan="3" rowspan="3"><a href="http://www.onroadtour.com.br/" target="_blank"><img src="/brmusic/Templates/web_22.jpg" alt="On Road Tour" width="163" height="114" border="0"></a></td>
<td><img src="espacio.gif" width="1" height="62" alt=""></td></tr><tr><td colspan="2"><img src="/brmusic/Templates/web_23.jpg" width="521" height="29" alt=""></td><td><img src="espacio.gif" width="1" height="29" alt=""></td></tr><tr><td><img src="/brmusic/Templates/web_24.jpg" width="489" height="23" alt=""></td><td colspan="2"><a href="http://www.estadologico.com/"><img src="/brmusic/Templates/web_25.jpg" alt="Diseño Web Granada -  Estado Lógico" width="109" height="23" border="0"></a></td>
<td><img src="espacio.gif" width="1" height="23" alt=""></td></tr><tr><td><img src="espacio.gif" width="14" height="1" alt=""></td><td><img src="espacio.gif" width="56" height="1" alt=""></td><td><img src="espacio.gif" width="70" height="1" alt=""></td><td><img src="espacio.gif" width="37" height="1" alt=""></td><td><img src="espacio.gif" width="489" height="1" alt=""></td><td><img src="espacio.gif" width="32" height="1" alt=""></td><td><img src="espacio.gif" width="77" height="1" alt=""></td><td></td></tr></table></div></body><!-- InstanceEnd --></html>
