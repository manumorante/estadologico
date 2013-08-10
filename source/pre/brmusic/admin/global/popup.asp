<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
	idioma = ""&request.QueryString("idioma")
	if idioma = "" then
		unerror = true : msgerror = "No se ha recibido el idioma."
	end if
	id = ""&request.QueryString("id")
	if id = "" then
		unerror = true : msgerror = "No se ha recibido la id."
	end if

	conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("../../datos/"& idioma &"/popup/popup.mdb")

	if not unerror then
		' Abro la conexion a la base de datos
		sql = "SELECT * FROM REGISTROS WHERE R_ID = " & id
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
		if not re.eof then
			nav_activo = cbool(re("R_ACTIVO"))
			nav_titulo = ""&re("R_TITULO")
			nav_fechaini = re("R_FECHAINI")
			nav_fechafin = re("R_FECHAFIN")
			nav_foto = ""&re("R_FOTO")
			nav_texto_arriba = ""&re("R_MEMO1")
			nav_texto_abajo = ""&re("R_MEMO2")
			nav_ancho = re("R_TEXT1")
			nav_alto = re("R_TEXT2")
			nav_tamanoigualfoto = ""&re("R_OPCION1")
			nav_enlace = ""&re("R_ENLACE")
			nav_color = ""&re("R_TEXT3")
		else
			unerror = true : msgerror = "No se ha encontrado ningún registro."
		end if
		re.close
		set re = nothing
	end if

%>

<html>
<head>
<title><%=nav_titulo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	<%if nav_color <> "" then%>
	background-color: <%=nav_color%>;
	<%end if%>
}
.Estilo1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9pt;
}
-->
</style>
<script>
function tamano(w,h) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
	moveTo(winl,wint)
	resizeTo(w+30,h+60)
}
</script>
</head>

<body class="bodyAdmin">
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><table border="0" cellspacing="0" cellpadding="0">
	<%if ""&nav_texto_arriba <> "" then%>
      <tr>
        <td><span class="Estilo1"><%=nav_texto_arriba%></span></td>
      </tr>
	<%end if%>
	
	<%if nav_foto <> "" then%>
      <tr>
        <td>
		<%if nav_enlace <> "" then%>
			<a href="<%=nav_enlace%>" target="_blank">
		<%end if%>
		<img src="../../datos/<%=idioma%>/popup/fotos/<%=nav_foto%>" <%if nav_tamanoigualfoto = "si" then%> onLoad="tamano(this.width,this.height)"<%end if%> border="0">
		<%if nav_enlace <> "" then%>
			</a>
		<%end if%>
		</td>
      </tr>
	  <%end if%>
	  
	<%if ""&nav_texto_abajo <> "" then%>
      <tr>
        <td><span class="Estilo1"><%=nav_texto_abajo%></span></td>
      </tr>
	<%end if%>
    </table>      </td>
  </tr>
</table>

</body>
</html>
