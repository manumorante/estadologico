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
        <div style="overflow:auto; width:521; height:440;">
          <p><span class="titulo-seccion">Lista de correo</span><br>
              <br>
            Inscríbete en nuestra lista de correo y recibe todas las novedades, próximos eventos, concursos, etc. Tan sólo rellena el siguiente  formulario:</p>
          <p>
            <%if request.Form() <> "" then
			%>
            <!--#include file="admin/inc_sendmail.asp" -->
            <%
			pFromAddress = "info@brmusic.net"
			pFromName = request.Form("nombre")
			pRecipient = "info@brmusic.net"
			pRecipientName = "Br Music"
			pSubject = "[brmusic.net] Lista de correo"

			pBody = pBody & "<font face=verdana size=2>"

			pBody = pBody & "<h3>Lista de correo</h3>"

			pBody = pBody & "<strong>Nombre / empresa:</strong> "		& request.Form("nombre")					&"<br>"
			pBody = pBody & "<strong>Ciudad:</strong> "					& request.Form("ciudad")					&"<br>"
			pBody = pBody & "<strong>Correo electrónico:</strong> "		& request.Form("email")						&"<br>"
			pBody = pBody & "<strong>Preferencias musicales:</strong> "	& request.Form("preferencias_musicales")	&"<br>"

			pBody = pBody & "</font>"

			enviar = sendMail(pFromAddress, pFromName, pRecipient, pRecipientName, pSubject, pBody)
			
			if enviar = "" then%>
            <strong>El formulario se ha enviado correctamente</strong>.<br>
  &iexcl;Ya formas parte de nuestra lista de correo!<br>
  <br>
  <strong>Muchas gracias</strong>
          <p><a href="lista_de_correo.asp">Volver</a> </p>
          <%else%>
          <strong>No se ha logrado enviar el formulario</strong>.<br>
Por favor, pruebe más tarde o escriba a info@brmusic.net.<br>
<br>
<font size="1">(<%=enviar%>)</font>
<br>
<br>

<strong>Muchas gracias</strong>
</p>
<p><a href="lista_de_correo.asp">Volver</a> </p>
<%end if
		else%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <form name="f" action="lista_de_correo.asp" method="post">
    <tr>
      <td width="25">&nbsp;</td>
      <td><table border="0" cellpadding="1" cellspacing="0">
        <tr>
          <td>Nombre / empresa:</td>
          <td colspan="3"><input name="nombre" type="text" class="campo" id="nombre" size="40"></td>
        </tr>
        <tr>
          <td>Ciudad:</td>
          <td><input name="ciudad" type="text" class="campo" id="ciudad" size="15"></td>
          <td>País:</td>
          <td><input name="pais" type="text" class="campo" id="pais" size="15"></td>
        </tr>
        <tr>
          <td>Correo electrónico:</td>
          <td colspan="3"><input name="email" type="text" class="campo" id="email" size="40"></td>
        </tr>
        <tr>
          <td>Preferencias musicales:</td>
          <td colspan="3"><input name="preferencias_musicales" type="text" class="campo" id="preferencias_musicales" size="40"></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td align="right">&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td align="right"><input type="image" src="arch/enviar.gif" alt="Enviar">
          </td>
        </tr>
      </table></td>
      <td>&nbsp;</td>
    </tr>
  </form>
</table>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="8">
  <tr>
    <td><table width="100%" border="1" cellpadding="8" cellspacing="0" bordercolor="#6C0A07" bgcolor="#F9F7E6">
      <tr>
        <td><p>Aquí teneis la lista de ganadores de los 10 lotes de productos con   material promocional. Espero sepais disculparnos por el retraso en el   anuncio. &iexcl;Gracias a todos los participantes!</p>
              <ol>
                <li> <strong>Sonia Slater</strong> - Torrevieja (Alicante)</li>
                <li> <strong>Oswaldo</strong> - Madrid</li>
                <li><strong>David Baeza</strong> - Algeciras (Cádiz)</li>
                <li><strong>Sergio Gonzalez</strong> - Estepona (Málaga)</li>
                <li><strong>Isaac Mora Pérez</strong> - Sevilla</li>
                <li><strong>Antonio Cabello</strong> - Motril (Granada)</li>
                <li><strong>Zoe Ball </strong>- Gibraltar</li>
                <li><strong>Sergio Portellano Puche</strong> - Atarfe (Granada)</li>
                <li><strong>Alain Cabessa Conesa</strong> - Málaga</li>
                <li><strong>Miriam Morales</strong> - Salobreña (Granada) </li>
              </ol></td>
      </tr>
    </table></td>
  </tr>
</table>
        </div>
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
