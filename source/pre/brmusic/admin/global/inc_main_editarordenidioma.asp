<%	nombre = request.Form("nombre")
	largo = len(nombre)
	if nombre <>"" then
		if largo > config_maxcarseccion then
		%><script language="javascript" type="text/javascript">
			alert(" * Error en longitud de nombre *\n\nEl nombre de sección escrito tiene <%=largo%> caracteres.\nPor favor escriba un nombre igual o inferior a <%=config_maxcarseccion%> caracteres.\n")
			location.href="main.asp?ac=adminsecciones"
		</script><%
		else
			id = request.Form("id")
			valor = request.Form("nombre")
			sql = "SELECT OI_TITULO,OI_ID FROM ORDENIDIOMA WHERE OI_ID = " & id
			consultaXOpen sql,2
				if not re.eof then
					re("OI_TITULO") = valor
					re.update()
				end if
			consultaXClose()
			
			Response.Redirect("main.asp?ac=adminordenidioma&msg=Sección editada correctamente.")
		end if


	else
		id = request.QueryString("id")
		if id <> "" and isNumeric(id) then
	
		sql = "SELECT * FROM ORDENIDIOMA WHERE OI_ID = " & id
		consultaXOpen sql,1
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
    <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Editar
          sub secci&oacute;n</font></b></td>
    <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
  </tr>
</table>
<br>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>Escriba el nuevo nombre para la secci&oacute;n <b><%=re("OI_TITULO")%></b> y pulse en el bot&oacute;n aceptar.<br>
      Recuerde que el nombre debe ser de <%=config_maxcarseccion%> car&aacute;teres como m&aacute;ximo. </td>
  </tr>
</table>
<br>
<table border="0" cellpadding="2" cellspacing="0">
  <tr>
    <td align="right">Nombre de la secci&oacute;n:      </td>
    <td><input name="nombre" type="text" class="campoAdmin" id="nombre" value="<%=re("OI_TITULO")%>" maxlength="100"></td>
  </tr>
</table>
<br>

<script>
f.id.value=<%=id%>
</script>
<table width="100%" border="0" cellpadding="2" cellspacing="0">
  <tr>
    <td align="right"><table border="0" cellspacing="0" cellpadding="1">
      <tr>
        <td><table border="0" cellpadding="2" cellspacing="0" onClick="location.href='main.asp?ac=adminordenidioma'" class="botonAdmin">
          <tr>
            <td>Volver</td>
          </tr>
        </table></td>
        <td><table border="0" cellpadding="2" cellspacing="0" onClick="f.submit()" class="botonAdmin">
          <tr>
            <td>Aceptar</td>
          </tr>
        </table></td>
      </tr>
    </table>      </td>
    </tr>
</table>
<br>
<br>
<%
			consultaXClose()
		else
			Response.Redirect("inicio.asp")
		end if
	end if
	%>