<!--#include file="inc/inc_conn.asp" -->
<!--#include file="inc/inc_rutinas.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="estilos.css" type="text/css" rel="stylesheet">
<title>Contactos</title>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
Dim key
key = ""& request.Form("key")
%>
<script language="javascript" type="text/javascript">
	ids = "|";
	function clickCheck(c) {
		if (c.checked) {
			ids = ids + c.value +"|";
		} else {
			ids = ids.replace("|"+ c.value +"|","|");
		}
		if (ids == "|") ids = "";
		top.f.e_contactos.value = ids;
		
	}
</script>
<form name="f" method="post" action="#">
<input name="imageField" type="image" src="arch/lupa.gif" width="18" height="18" border="0" align="absmiddle">
<input name="key" type="text" class="campo" value="<%=key%>" size="15">
<img src="arch/linea.gif" width="100%" height="1">
<%
sql = "SELECT * FROM CONTACTOS ORDER BY C_NOMBRE"
if key <> "" then
	sql = sql & " WHERE C_NOMBRE LIKE '%"& key &"%'"
	sql = sql & " OR C_ID = "& mNumero(key)
end if
set re = mConsulta(sql,conn_,2)%>
	<table  border="0" cellspacing="0" cellpadding="1">
		<%while not re.eof
			fila1 = "#EDF3FE"
			fila2 = "#FFFFFF"
			if fila = fila1 then
				fila = fila2
			else
				fila = fila1
			end if%>
			<tr bgcolor="<%=fila%>">
			<td><input name="id_contacto" type="checkbox" id="id_contacto<%=re("C_ID")%>" title="<%=re("C_ID")%>" onClick="clickCheck(this);" value="<%=re("C_ID")%>">			  <label for="id_contacto<%=re("C_ID")%>"><%=re("C_NOMBRE")%>&nbsp;<%=re("C_APELLIDOS")%></label>&nbsp;</td>
			</tr>
			<%re.movenext
		wend%>
</table>
</form>
</body>
</html>
