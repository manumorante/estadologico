<%idiactual = ""&request.Form("idiactual")
	if idiactual = "" then
		idiactual = "esp"
	end if

	' Insertar una sección nueva si se recibe el nombre al enviar un formulario
	titulo = ""&replace(request.Form("nuevaseccion"),"'","")
	titulo = replace(titulo,chr(34),"") ' quito las comillas dobles
	largo = len(titulo)
	if ""&titulo <> "" then
		if largo > config_maxcarseccion then
			%><script language="javascript" type="text/javascript">alert(" * Error en longitud de nombre *\n\nEl nombre de sección escrito tiene <%=largo%> caracteres.\nPor favor escriba un nombre igual o inferior a <%=config_maxcarseccion%> caracteres.\n")</script><%
		else
			' Busco si hay otra sección con el mismo nombre
			consultaXOpen "SELECT OI_TITULO FROM ORDENIDIOMA WHERE OI_TITULO = '"& titulo &"'",1
			if not re.eof then
				consultaXClose()
				%>
				<script>
				f.ac.value = "adminordenidioma"
				f.msgerror.value = "Hay una sección con el mismo nombre."
				f.submit()
				</script>
				<%
			else
				if request.Form("pos") = "inicio" then
					n = 0.9
				else
					n = "9999"
				end if
			
				' Hago la inserción
				sql = "INSERT INTO ORDENIDIOMA (OI_TITULO, OI_ORDEN_ESP, OI_ORDEN_ENG, OI_ORDEN_FRA, OI_ORDEN_DEU, OI_ORDEN_ITA) VALUES ('"& titulo &"','"& n &"','"& n &"','"& n &"','"& n &"','"& n &"')"
				call exeSql(sql, conn_)
				
				reordenaOrdenIdioma()
				%>
				<script>
				try{
					//var f1 = parent.frames[0].f // Frame de la izquierda
					//f1.ac.value = ""
					//f1.action = "main.asp"
					//f1.target = ""
					//f1.submit()
				}catch(unerror){}
				</script>
				<%
			end if
		end if ' largo 
	end if
	

	sql = "SELECT *"
	sql = sql & " FROM ORDENIDIOMA"
	sql = sql & " ORDER BY OI_ORDEN_"& ucase(idiactual)

	s
	
	if not unerror then
%>
	<script language="javascript" type="text/javascript">
		<!--
		function visible(id,idi,v){
			f.ac.value = "ordenidiomavisible"
			f.idi.value = idi
			f.v.value = v
			f.id.value = id
			f.submit()
		}
		//-->
	</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Orden por idioma</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>      Pulse sobre el idioma deseado para ordenar por el.<br>
      Los cambios sobre el orden s&oacute;lo afectar&aacute;n al idioma seleccionado (columna coloreada).<br>
      Recuerde que el nombre debe ser de <%=config_maxcarseccion%> car&aacute;teres como m&aacute;ximo. </td>
  </tr>
</table>
<br>
<table  border="0" cellpadding="1" cellspacing="0">
  <tr valign="middle">
    <td align="right"><input name="nuevaseccion" type="text" class="campoAdmin" id="nuevaseccion" maxlength="100">      </td>
    <td><input type="submit" class="botonAdmin" value="Enviar"></td>
    <td>
	<input name="pos" id="pos_primera" type="radio" value="inicio" <%if request.QueryString("pos") = "inicio" or ""&request.QueryString("pos") = "" then response.Write "checked" end if%>><label for="pos_primera">Primera</label>
	<input name="pos" id="pos_ultima" type="radio" value="fin" <%if request.QueryString("pos") = "fin" then response.Write "checked" end if%>><label for="pos_ultima">&Uacute;ltima</label></td>
  </tr>
</table>
	<br>
	<%if request.QueryString("msgerror")&request.form("msgerror") <> "" then%>
	  <table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td align="center" class="Error"><b>ATENCIÓN:</b> <%=request.QueryString("msgerror")&request.form("msgerror")%></td>
  </tr>
</table>
	<%end if%>

	
	<%if request.QueryString("msgerror") <> "" or request.QueryString("msg") <> "" then%>
	<script>
			try{
				var f = parent.frames[0].f // Frame de la izquierda
				f.ac.value = ""
				f.action = "main.asp"
				f.target = ""
				f.submit()
			}catch(unerror){
//
			}
	</script>
	<%end if%>
	
<%if re.eof then%>
	<div align="center"><b>No hay ninguna secci&oacute;n.</b></div>
<%else%>
<table width="100%"  border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td colspan="8">Orden por: <span class="tituloazonaAdmin"><b>

	<%select case idiactual
	case "esp"%>
	Espa&ntilde;ol
	<%case "eng"%>
	Ingl&eacute;s
	<%case "fra"%>
	Franc&eacute;s
	<%case "deu"%>
	Alem&aacute;n
	<%case "ita"%>
	Italiano
	<%end select%>

    </b></span>
	<input name="idiactual" type="hidden" value="<%=idiactual%>">
	<input name="idi" type="hidden" value="">
	<input name="v" type="hidden" value=""></td>
    </tr>
  <tr class="fondoAdmin">
    <td align="center"><b>Pos.</b></td>
    <td width="100%"><b>Nombre</b></td>
    <td align="center"><a href="JavaScript:chanIdi('esp')" class="aAdmin"><b>Espa&ntilde;ol</b></a></td>
    <td align="center"><a href="JavaScript:chanIdi('eng')" class="aAdmin"><b>Ingl&eacute;s</b></a></td>
    <td align="center"><a href="JavaScript:chanIdi('fra')" class="aAdmin"><b>Franc&eacute;s</b></a></td>
    <td align="center"><a href="JavaScript:chanIdi('deu')" class="aAdmin"><b>Alem&aacute;n</b></a></td>
    <td align="center"><a href="JavaScript:chanIdi('ita')" class="aAdmin"><b>Italiano</b></a></td>
    <td bgcolor="#f7f7f7">&nbsp;</td>
  </tr>
  <%while not re.eof%>
  <tr>
    <td align="center" bgcolor="#FFFFFF" class="fondoAdmin"><%=re("OI_ORDEN_"& idiactual)%></td>
    <td bgcolor="#FFFFFF"><%=re("OI_TITULO")%></td>
    <td align="center" <%if idiactual = "esp" then Response.Write "class='CeldaActiva'" else Response.Write "class='CeldaInactiva'" end if%>><%if re("OI_VISIBLE_ESP") then%><a href="JavaScript:visible(<%=re("OI_ID")%>,'esp',0)"><img src="img/ojo_activo.gif" alt=" Ocultar " width="18" height="18" border="0"></a><%else%><a href="JavaScript:visible(<%=re("OI_ID")%>,'esp',1)"><img src="img/ojo_inactivo.gif" alt=" Mostrar " width="18" height="18" border="0"></a><%end if%></td>
    <td align="center" <%if idiactual = "eng" then Response.Write "class='CeldaActiva'" else Response.Write "class='CeldaInactiva'" end if%>><%if re("OI_VISIBLE_ENG") then%><a href="JavaScript:visible(<%=re("OI_ID")%>,'eng',0)"><img src="img/ojo_activo.gif" alt=" Ocultar " width="18" height="18" border="0"></a><%else%><a href="JavaScript:visible(<%=re("OI_ID")%>,'eng',1)"><img src="img/ojo_inactivo.gif" alt=" Mostrar " width="18" height="18" border="0"></a><%end if%></td>
    <td align="center" <%if idiactual = "fra" then Response.Write "class='CeldaActiva'" else Response.Write "class='CeldaInactiva'" end if%>><%if re("OI_VISIBLE_FRA") then%><a href="JavaScript:visible(<%=re("OI_ID")%>,'fra',0)"><img src="img/ojo_activo.gif" alt=" Ocultar " width="18" height="18" border="0"></a><%else%><a href="JavaScript:visible(<%=re("OI_ID")%>,'fra',1)"><img src="img/ojo_inactivo.gif" alt=" Mostrar " width="18" height="18" border="0"></a><%end if%></td>
    <td align="center" <%if idiactual = "deu" then Response.Write "class='CeldaActiva'" else Response.Write "class='CeldaInactiva'" end if%>><%if re("OI_VISIBLE_DEU") then%><a href="JavaScript:visible(<%=re("OI_ID")%>,'deu',0)"><img src="img/ojo_activo.gif" alt=" Ocultar " width="18" height="18" border="0"></a><%else%><a href="JavaScript:visible(<%=re("OI_ID")%>,'deu',1)"><img src="img/ojo_inactivo.gif" alt=" Mostrar " width="18" height="18" border="0"></a><%end if%></td>
    <td align="center" <%if idiactual = "ita" then Response.Write "class='CeldaActiva'" else Response.Write "class='CeldaInactiva'" end if%>><%if re("OI_VISIBLE_ITA") then%><a href="JavaScript:visible(<%=re("OI_ID")%>,'ita',0)"><img src="img/ojo_activo.gif" alt=" Ocultar " width="18" height="18" border="0"></a><%else%><a href="JavaScript:visible(<%=re("OI_ID")%>,'ita',1)"><img src="img/ojo_inactivo.gif" alt=" Mostrar " width="18" height="18" border="0"></a><%end if%></td>
    <td align="right"><table border="0" cellspacing="0" cellpadding="0">
      <tr>
	  <%if NOT re("OI_BLOQUEADA") then%>
	  	<td><a href="javascript:moverOrdenIdioma(<%=re("OI_ID")%>,'subir')"><img src="img/flecha_arriba_h.gif" alt=" Subir " width="15" height="18" border="0"></a><a href="javascript:moverOrdenIdioma(<%=re("OI_ID")%>,'bajar')"><img src="img/flecha_abajo_h.gif" alt=" Bajar " width="15" height="18" border="0"></a></td>
		<%end if%>
	  <%if re("OI_RENOMBRAR") then%>
	  	<td><a href="javascript:editarOrdenIdioma(<%=re("OI_ID")%>)"><img src="img/lapiz.gif" alt=" Editar " width="18" height="18" border="0"></a></td>
	  <%end if%>
	  <%if re("OI_ELIMINAR") and re("OI_REGISTROS") =0 then%>
        <td><a href="javascript:eliminarOrdenIdioma(<%=re("OI_ID")%>)"><img src="img/papelera.gif" alt=" Eliminar " width="18" height="18" border="0"></a></td>
	  <%end if%>
      </tr>
    </table>      </td>
  </tr>
  <%re.movenext : wend%>
</table>

<script language="javascript" type="text/javascript">
<!--
	function chanIdi(idi){
		f.idiactual.value = idi
		f.submit()
	}
//-->
</script>

<%end if ' re.eof%>
<br>

<%	consultaXClose()
end if%>
<script>f.ac.value = "adminordenidioma"</script>