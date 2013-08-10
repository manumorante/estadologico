<html><!-- InstanceBegin template="/Templates/base2_eng.dwt.asp" codeOutsideHTMLIsLocked="false" --><head><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /><!-- InstanceBeginEditable name="doctitle" -->
<title>BR Music International .:. festivales de musica, nacionales e internacionales, escenarios, camerinos, vallas, catering, montaje fijaci&oacute;n de carteleria, promoci&oacute;n y mailing, grupos electr&oacute;genos</title>
<!-- InstanceEndEditable --><link rel="stylesheet" type="text/css" href="arch/estilos.css"><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable --></head><body bgcolor="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"><div align="center"><table id="Tabla_01" width="776" height="580" border="0" cellpadding="0" cellspacing="0"><tr><td colspan="6"><img src="/brmusic/Templates/web_eng_01.jpg" width="775" height="22" alt=""></td><td><img src="espacio.gif" width="1" height="22" alt=""></td></tr><tr><td rowspan="16"><img src="/brmusic/Templates/web_eng_02.jpg" width="14" height="557" alt=""></td><td rowspan="2"><a href="/index_eng.asp"><img src="/brmusic/Templates/web_eng_03.jpg" alt="" width="126" height="111" border="0"></a></td>
<td colspan="4"><img src="/brmusic/Templates/web_eng_04.jpg" width="635" height="65" alt=""></td><td><img src="espacio.gif" width="1" height="65" alt=""></td></tr><tr><td rowspan="11"><img src="/brmusic/Templates/web_eng_05.jpg" width="37" height="326" alt=""></td><td colspan="2" rowspan="13" align="left" valign="top" background="/brmusic/Templates/web_eng_06.jpg" bgcolor="#EDE6C9"><!-- InstanceBeginEditable name="Cuerpo" -->
        <div style="overflow:auto; width:521; height:440;">
		<p><span class="titulo-seccion">Mailing list</span><br>  
          <br>
        Sign up for our mailing list and you will be informed of all news, upcoming gigs, contests, etc.
It is just as easy as filling in the form provided below:</p>
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
        <p><a href="lista_de_correo.asp">Volver</a>        </p>
		      <%else%>
				<strong>No se ha logrado enviar el formulario</strong>.<br>
				Por favor, pruebe m&aacute;s tarde o escriba a info@brmusic.net.<br>
				<br>
		<strong>Muchas gracias</strong></p>
        <p><a href="lista_de_correo.asp">Volver</a>        </p>
              <%end if
		else%>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
		<form name="f" action="lista_de_correo.asp" method="post">
          <tr>
            <td width="25">&nbsp;</td>
            <td><table border="0" cellpadding="1" cellspacing="0">
              <tr>
                <td>Name/Company:</td>
                <td colspan="3"><input name="nombre" type="text" class="campo" id="nombre" size="48"></td>
                </tr>
              <tr>
                <td>City:</td>
                <td><input name="ciudad" type="text" class="campo" id="ciudad" size="15"></td>
                <td>Country:</td>
                <td><input name="pais" type="text" class="campo" id="pais" size="15"></td>
              </tr>
              <tr>
                <td>Email:</td>
                <td colspan="3"><input name="email" type="text" class="campo" id="email" size="48"></td>
                </tr>
              <tr>
                <td>Music preferences:</td>
                <td colspan="3"><input name="preferencias_musicales" type="text" class="campo" id="preferencias_musicales" size="48"></td>
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
                <td align="right"><input type="image" src="arch/enviar_eng.gif" alt="Enviar">                  </td>
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
                <td><p>It finally happened!The ten winners of our&nbsp; memorabilia boxes are:</p>
                  <ol>
                    <li> <strong>Sonia Slater</strong> - Torrevieja (Alicante)</li>
                    <li> <strong>Oswaldo</strong> - Madrid</li>
                    <li><strong>David Baeza</strong> - Algeciras (C&aacute;diz)</li>
                    <li><strong>Sergio Gonzalez</strong> - Estepona (M&aacute;laga)</li>
                    <li><strong>Isaac Mora P&eacute;rez</strong> - Sevilla</li>
                    <li><strong>Antonio Cabello</strong> - Motril (Granada)</li>
                    <li><strong>Zoe Ball </strong>- Gibraltar</li>
                    <li><strong>Sergio Portellano Puche</strong> - Atarfe (Granada)</li>
                    <li><strong>Alain Cabessa Conesa</strong> - M&aacute;laga</li>
                    <li><strong>Miriam Morales</strong> - Salobre&ntilde;a (Granada) </li>
                  </ol></td>
              </tr>
            </table></td>
          </tr>
        </table>
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
