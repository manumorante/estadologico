<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<title>BR Music International .:. festivales de musica, nacionales e internacionales, escenarios, camerinos, vallas, catering, montaje fijación de carteleria, promoción y mailing, grupos electrógenos</title>

<style type="text/css">
<!--
A.blanco:visited {
	color:#ffffff;
} 
td {
	font-family: Tahoma;
	color: #594700;
	font-size: 14px;
}
h3 {
	color: #9E4837;
	font-size: 24px;
	font-weight: normal;
}
A.blanco:active {
	color:#ffffff;
} 
A.blanco:link {
	color:#ffffff;
} 
A.blanco:hover {
	color: #FFFFFF;
}
.datos-concierto {
	color: #9E4837;
	font-weight: bold;
	font-size: 10px;
}
.titulo-seccion {
	font-size: 20px;
	color: #630708;
}
-->
</style>
</head>
<body bgcolor="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	url = request.ServerVariables("URL")
	Response.Write url
%>
<div align="center">
  <table id="Tabla_01" width="776" height="605" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td colspan="5"><img src="Templates/web_01.jpg" width="775" height="21" alt=""></td>
      <td><img src="Templates/espacio.gif" width="1" height="21" alt=""></td>
    </tr>
    <tr>
      <td rowspan="13"><img src="Templates/web_02.jpg" width="14" height="584" alt=""></td>
      <td rowspan="2"><a href="../"><img src="Templates/web_03.jpg" alt="" width="126" height="111" border="0"></a></td>
      <td colspan="3"><img src="Templates/web_04.jpg" width="635" height="78" alt=""></td>
      <td><img src="Templates/espacio.gif" width="1" height="78" alt=""></td>
    </tr>
    <tr>
      <td rowspan="9"><img src="Templates/web_05.jpg" width="37" height="326" alt=""></td>
      <td width="521" height="440" rowspan="11" align="left" valign="top" bordercolor="#F1EACD" background="Templates/web_06.jpg" bgcolor="#F1EACD"><%idioma = "esp" : cualid = "conciertos"%>
        <!--#include file="visores/inc_conn.asp" -->
        <%

	if not unerror then
		sql = "SELECT * FROM REGISTROS ORDER BY R_FECHA"
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
                  <%if re("R_TEXT1") <> "" then%>
                  <strong><%=re("R_TEXT1")%></strong><br>
                  <%end if%>
                  <%if ""& re("R_FECHA") <> "0:00:00" and ""& re("R_FECHA") <> "" then%>
                  <span class="datos-concierto">DíA:</span> <%=re("R_FECHA")%><br>
                  <%end if%>
                  <%if re("R_TEXT3") <> "" then%>
                  <span class="datos-concierto">RECINTO:</span> <%=re("R_TEXT3")%><br>
                  <%end if%>
                  <%if re("R_TEXT4") <> "" then%>
                  <span class="datos-concierto">LUGAR:</span> <%=re("R_TEXT4")%><br></td>
              <%end if%>
              <td align="right"><%if re("R_ICONO") <> "" then%>
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
        <!--#include file="inc_alerta.asp" --></td>
      <td rowspan="12"><img src="Templates/web_07.jpg" width="77" height="506" alt=""></td>
      <td><img src="Templates/espacio.gif" width="1" height="33" alt=""></td>
    </tr>
    <tr>
      <td><img src="Templates/web_08.jpg" width="126" height="37" alt=""></td>
      <td><img src="Templates/espacio.gif" width="1" height="37" alt=""></td>
    </tr>

		<tr>
		<td><a href="index.asp"><img src="Templates/web_09.jpg" alt="Conciertos" width="126" height="20" border="0"></a></td>
		<td><img src="Templates/espacio.gif" width="1" height="20" alt=""></td>
		</tr>
    <tr>
      <td><img src="Templates/web_10.jpg" alt="Festivales" width="126" height="21" border="0"></td>
      <td><img src="Templates/espacio.gif" width="1" height="21" alt=""></td>
    </tr>

		<tr>
		  <td><img src="Templates/web_11.jpg" alt="Entradas" width="126" height="22" border="0"></td>
		  <td><img src="Templates/espacio.gif" width="1" height="22" alt=""></td>
		</tr>


		<tr>
		  <td><img src="Templates/web_12.jpg" alt="Servicios" width="126" height="22" border="0"></td>
		  <td><img src="Templates/espacio.gif" width="1" height="22" alt=""></td>
		</tr>


		<tr>
		  <td><img src="Templates/web_13.jpg" alt="Prensa" width="126" height="21" border="0"></td>
		  <td><img src="Templates/espacio.gif" width="1" height="21" alt=""></td>
		</tr>


		<tr>
		  <td><img src="Templates/web_14.jpg" alt="Contacto" width="126" height="22" border="0"></td>
		  <td><img src="Templates/espacio.gif" width="1" height="22" alt=""></td>
		</tr>


    <tr>
      <td><img src="Templates/web_15.jpg" width="126" height="128" alt=""></td>
      <td><img src="Templates/espacio.gif" width="1" height="128" alt=""></td>
    </tr>
    <tr>
      <td colspan="2"><a href="http://www.brmusic.net/Festival_Atarfe_Vega_Rock/" target="_blank"><img src="Templates/web_16.jpg" alt="Festival Atarfe Vega Rock" width="163" height="52" border="0"></a></td>
      <td><img src="Templates/espacio.gif" width="1" height="52" alt=""></td>
    </tr>
    <tr>
      <td colspan="2" rowspan="2"><img src="Templates/web_17.jpg" width="163" height="128" alt=""></td>
      <td><img src="Templates/espacio.gif" width="1" height="62" alt=""></td>
    </tr>
    <tr>
      <td><img src="Templates/web_18.jpg" width="521" height="66" alt=""></td>
      <td><img src="Templates/espacio.gif" width="1" height="66" alt=""></td>
    </tr>
  </table>
  <table width="776" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="right"><a href="http://www.estadologico.com/" target="_blank"><img src="arch/diseno_web.gif" alt="Diseño Web y Programación - Granada" width="91" height="31" hspace="12" border="0"></a></td>
    </tr>
  </table>
  </div>
</body>
</html>
