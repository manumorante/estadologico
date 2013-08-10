<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Server.ScriptTimeOut = 1000000000  ' 20 minutos

Dim conn_
conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("../../data/contactos.mdb")
%>
<html>
<head>
<title>Contactos</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9pt;
}
body {
	background-color: #ECE9D8;
}
.campo {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9pt;
	padding-left: 4px;
	padding-right: 4px;
}
.campotype {
	font-family: "Courier New", Courier, mono;
}
-->
</style></head>

<body>
ZONA DE MAILING 
<hr size="1">
<a href="Default.asp?ac=redactar%20">Redactar</a> <a href="Default.asp?ac=organizar">Organizar</a> <a href="purgar.asp">Purgar</a> <a href="index.asp">Agregar</a><br>
<br>
<%
ac = ""&request.QueryString("ac")
select case ac
case "redactar"%>

<table width=""  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td bgcolor="#FFFFFF">Inicio / <b>Redactar</b></td>
  </tr>
</table>
<br>
<script>
	function envio() {
		if(confirm("¿Seguro que desea enviar el mensaje?")){
			f.botonenviar.disabled = true
			return true
		} else {
			return false
		}
	}
	
</script>
<form name="f" action="Default.asp?ac=envio" method="post" onSubmit="return envio()">
<table width="100%"  border="0" cellpadding="2" cellspacing="0" id="tabla">
  <tr>
    <td width="14%" align="right">Nombre: </td>
    <td width="86%"><input name="minombre" type="text" class="campo" id="minombre" value="Tradepunk.com"></td>
  </tr>
  <tr>
    <td align="right">Email: </td>
    <td><input name="miemail" type="text" class="campo" id="miemail" value="tradepunk@estadologico.com"></td>
  </tr>
  <tr>
    <td align="right">Asunto: </td>
    <td><input name="asunto" type="text" class="campo" id="asunto" value="Esto es una prueba"></td>
  </tr>
  <tr>
    <td colspan="2"><textarea name="cuerpo" cols="" rows="10" wrap="virtual" class="campo" id="cuerpo" style="width:100%;">Prueba</textarea></td>
  </tr>
  <tr>
    <td colspan="2" align="right"><input name="prueba" type="checkbox" value="1" checked>
      Enviar s&oacute;lo una prueba.
        <input name="botonenviar" type="submit" id="botonenviar" value="Enviar"></td>
  </tr>
</table>
</form>
<%case "envio"%>

	<table width=""  border="0" cellspacing="0" cellpadding="2">
	  <tr>
		<td bgcolor="#FFFFFF">Inicio / Redactar / <b>Enviar</b></td>
	  </tr>
	</table>
	<br>

<%
Function enviarEmail(minombre,miemail,destino,asunto,cuerpo)
	on error resume next
	Set Mail = Server.CreateObject("Persits.MailSender")
	Mail.Host = "smtp.estadologico.com"
	Mail.From = miemail ' Dirección del que envia
	Mail.FromName = minombre ' Nombre del que envia

	' Destino (to)
	if inStr(destino,",")>0 then
		emails = split(destino,",")
		for each email in emails
			Mail.AddAddress email
		next
	else
		Mail.AddAddress destino
	end if

	' Asunto (subject)
	Mail.Subject = asunto

	' Cuerpo (body)
	Mail.Body = cuerpo
	' Html
	Mail.IsHTML = True 
	
	bSuccess = False
		Mail.Send
		If Err <> 0 Then
			enviarEmail = Err.description
		else
			enviarEmail = true
	  End If
	 on error goto 0

end function

	minombre = request.Form("minombre")
	miemail = request.Form("miemail")
	asunto = request.Form("asunto")
	cuerpo = request.Form("cuerpo")

	if request.Form("prueba") = 1 then
		destino = "estadologico@hotmail.com"
'		destino = "saviabruta@hotmail.com"
		envio = enviarEmail(minombre,miemail,destino,asunto,cuerpo)
		if envio <> true then
			Response.Write "ERROR:<br>" & envio
		else
			Response.Write "<b>Se ha enviado un E-mail a '"& destino &"'</b>."
		end if
	else
'		sql = "SELECT TOP 10 * FROM EMAILS WHERE E_EMAIL LIKE '%manolo%' ORDER BY E_EMAIL"
		sql = "SELECT * FROM EMAILS ORDER BY E_ID"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_ : re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
		numEnviados = 0
		for n=1 to re.recordcount
			destino = trim(re("E_EMAIL"))
			envio = enviarEmail(minombre,miemail,destino,asunto,cuerpo)
			if envio <> true then
				msgerror = msgerror & "<b>"& destino & "</b>: " & envio & "<br>"
			else
				enviados = enviados & destino &"<br>"
				numEnviados = numEnviados + 1
			end if
			re.movenext()
		next

		destino = "estadologico@hotmail.com"
		envio = enviarEmail(minombre,miemail,destino,asunto,cuerpo)
		if envio <> true then
			msgerror = msgerror & "<b>"& destino & "</b>: " & envio & "<br>"
		else
			enviados = enviados & destino &"<br>"
		end if
		
		if msgerror <> "" then
			Response.Write "FALLIDOS:<br>" & msgerror & "<br><br>"
		else
			Response.Write "NINGÚN FALLO:<br><br>"
		end if

		if enviados <> "" then
			Response.Write "ENVIADOS: ("& numEnviados &")<br>"& enviados &"<br><br>"
		else
			Response.Write "NINGÚN ENVIO:<br><br>"
		end if
	end if

	' -----------  PRUEBAS
	'destino = "estadologico@hotmail.com, manolo@estadologico.com, info@estadologico.com, info@uvemoda.com, info@sellourbano.com"
	' -----------





case "borrar"
	id = request.QueryString("id")
	if ""&id<> "" and isNumeric(id) then
		sql = "DELETE FROM EMAILS WHERE E_ID = "& id
		set re = Server.CreateObject("ADODB.Connection")
		re.Open conn_ : re.execute sql : re.close() : set re = nothing
	end if
	Response.Redirect("default.asp?ac=organizar")
case "organizar"
	sql = "SELECT * FROM EMAILS"
	
	key = request.QueryString("key")
	if key <> "" then
		sql = sql & " AND E_EMAIL LIKE '%"&key&"%'"
	end if
	
	sql = sql & " ORDER BY E_EMAIL"
	set re = Server.CreateObject("ADODB.Recordset")
	re.ActiveConnection = conn_ : re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
%>
<script>
function borrarid(id,email){
	if (confirm("¿Borrar "+email+"?")){
		location.href="default.asp?ac=borrar&id="+ id +""
	}
}
function buscar(){
	location.href="default.asp?ac=organizar&key="+f.c1.value
	return false
}
</script>
<table width=""  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td bgcolor="#FFFFFF">Inicio / <b>Organizar</b></td>
  </tr>
</table>
<br>
<form action="#" method="post" name="f" onSubmit="return buscar()">
<input name="c1" type="text" class="campo" id="c1" value="<%=request.QueryString("key")%>" size="30" maxlength="255">
<input type="submit" class="campo" value="buscar!">
</form>
<%
total = re.recordcount
if isNumeric(request.QueryString("pag")) and ""&request.QueryString("pag") <> "" then
	pag = 0+request.QueryString("pag")
else
	pag = 30
end if

ini = request.QueryString("ini")
if ""&ini = "" then
	ini = 0
end if

if ini <0 then
	ini = 0
end if

if ini >= pag then
	ant = ini -pag
else
	ant = 0
end if

if ant < 0 then
	ant = 0
end if

re.move ini
%>
<span class="campotype"><a href="Default.asp?ac=organizar&ini=<%=ant%>&pag=<%=pag%>">&lt;&lt; Anterior</a> --------------------------[<%=ini%>/<%=ini+pag%> de <%=total%>]------------------------------ <a href="Default.asp?ac=organizar&ini=<%=ini+pag%>&pag=<%=pag%>">Siguiente &gt;&gt; </a></span><br>
<%for n=1 to pag
	if not re.eof then
		num = ini + n
		email = trim(re("E_EMAIL"))%><span class="campotype"><%="---   "&left (re("E_EMAIL") &" -----------------------------------------------------------------",75)&right("000000"&num,4)&" - "&right("000000"&re("E_ID"),6) %></span> <a href="JavaScript:borrarid(<%=re("E_ID")%>,'<%=re("E_EMAIL")%>')">borrar</a><br>
	<%end if
	re.movenext
next

	on error resume next
	re.Close()
	set re = nothing
	on error goto 0
%>

<span class="campotype"><a href="Default.asp?ac=organizar&ini=<%=ant%>&pag=<%=pag%>">&lt;&lt; Anterior</a> --------------------------[<%=ini%>/<%=ini+pag%> de <%=total%>]------------------------------ <a href="Default.asp?ac=organizar&ini=<%=ini+pag%>&pag=<%=pag%>">Siguiente &gt;&gt; </a></span><br>
Pag. <a href="Default.asp?ac=organizar&ini=<%=ini%>&pag=30">30</a> - <a href="Default.asp?ac=organizar&ini=<%=ini%>&pag=50">50</a> - <a href="Default.asp?ac=organizar&ini=<%=ini%>&pag=100">100</a> - <a href="Default.asp?ac=organizar&ini=<%=ini%>&pag=300">300</a> - <a href="Default.asp?ac=organizar&ini=<%=ini%>&pag=500">500</a> - <a href="Default.asp?ac=organizar&ini=<%=ini%>&pag=1000">1000</a><%case else%>
	Bienvenid@. Elije una opción.

<%end select%>

<br>

<br>
<hr size="1">
</body>
</html>
