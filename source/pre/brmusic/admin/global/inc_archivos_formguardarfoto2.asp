<%
	id = numero(request.QueryString("id"))
	if id > 0 then
	
		' Abrir el registro...
		sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
		consultaXOpen sql,1
%>

		<script language="javascript" type="text/javascript">
			<!--
			function envio(){
				if (""+ f.archivo.value == ""){
					alert("Seleccione una foto.");
					return false;
				}
				f.enviar.disabled = true
				return true
			}
			//-->
		</script>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
	<td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar imagen (IE0)</font></b></td>
	<td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
	</tr>
</table>

<br />
<form action="archivos.asp?ac=opciones_foto&id=<%=id%>&icono=<%=request.QueryString("icono")%>" method="post" enctype="multipart/form-data" name="f" onSubmit="return envio()">
  <%if reTotal >0 then%>
  <table width="100%"  border="0" cellpadding="2" cellspacing="0">
    <tr>
      <td><b><%=config_nom_titulo%>: </b></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><%=re("R_TITULO")%></td>
    </tr>
  </table>
  <br />
  <%end if%>
  <!--
  <input name="archivo" type="file" id="archivo" />
  -->
  
  <!--
  <input name="" type="button" class="botonAdmin" onclick="window.history.back()" value="Cancelar" />
  <input name="enviar" type="submit" class="botonAdmin" id="enviar" value="Enviar" />
  -->

</form>

<%
	consultaXClose()
end if
%>