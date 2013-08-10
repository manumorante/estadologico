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
	<td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Insertar imagen</font></b></td>
	<td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
	</tr>
</table>

<form name="f" method="post" action="archivos.asp?ac=opciones_foto&id=<%=id%>&icono=<%=request.QueryString("icono")%>" onSubmit="return envio()">

<table width="101%"  border="0" cellspacing="0" cellpadding="2">
	<tr>
	<td colspan="2" align="center" valign="middle">
	<%if reTotal >0 then%>

		<table width="100%"  border="0" cellpadding="2" cellspacing="0">
			<tr>
			<td><b><%=config_nom_titulo%>: </b></td>
			</tr>
			<tr>
			<td bgcolor="#FFFFFF"><%=re("R_TITULO")%></td>
			</tr>
		</table>
		<br>

	<%end if%>

	<table border="0" cellspacing="0" cellpadding="0">
		<tr>
		<td><input name="archivo" type="text" class="campoAdmin" id="archivo" size="50" readonly="true"></td>
		<td>&nbsp;<input type="button" class="botonAdmin" onClick="parent.frames[1].f.archivo.click()" value="Examinar ..."></td>
		</tr>
	</table></td>
	</tr>
	<tr>
	<td colspan="2" align="right">&nbsp;</td>
	</tr>
	<tr>
	<td align="center">&nbsp;</td>
	<td align="right">
	<input name="" type="button" class="botonAdmin" onClick="window.history.back()" value="Cancelar">
	<input name="enviar" type="submit" class="botonAdmin" id="enviar" value="Enviar"></td>
	</tr>
</table>

</form>

<%
	consultaXClose()
end if
%>