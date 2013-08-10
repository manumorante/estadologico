<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="inc_lib_email.asp" -->
<%
	Dim conn_
	conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("../../data/contactos.mdb")
%>
<style type="text/css">
<!--
body {
	margin-left: 5px;
	margin-top: 5px;
	margin-right: 5px;
	margin-bottom: 5px;
	background-color: #ece9d8;
}
body,td,th {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 7.5pt;
}
.areacadena {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 7.5pt;
	color: #316AC5;
}
-->
</style>
<%
cadena = request.Form("cadena")
if cadena <> "" then
%>

<title>Lector de direcciones de email</title><br>
<br>
<table width="400" border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#F4F2E8">
	<tr valign="top">
	  <td align="right"><input name="Submit" type="submit" style="height:50px" onClick="location='index.asp'" value="Insertar mas"></td>
  </tr>
	<tr valign="top">
	  <td><strong>Recibidos:</strong> (<%=cuentaPalabras(emails,"@")%>)</td>
  </tr>
	<tr valign="top">

<%

	emails = getEmails(request.Form("cadena"))
	%>
	<td><%=replace(emails,"|","<br>")%></td>
	<%
	
	
	if 1=2 then
	' Creo una cadana con todos los emails malos
	sql = "SELECT EM_EMAIL FROM MALOS"
	set re = Server.CreateObject("ADODB.Recordset")
	malos = "|"
	re.ActiveConnection = conn_ : re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 2 : re.Open()
		while not re.eof
			malos = malos & re("EM_EMAIL") & "|"
			re.movenext
		wend
	re.close()
	set re = nothing
	
	' Compruebo que los emails que he obtenido no esten en la lista mala
	arr = split(emails,"|")
	buenos = "|"
	for each a in arr
		' Si no esta en la lista lo añado a la lista de buenos
		if inStr(malos,"|"& a &"|") <= 0 then
			buenos = buenos & a & "|"
		end if
	next
	end if
	%>
	<%
	if 1=2 then
	' Ahora insertotodos los bueno en su tabla
	sql = "SELECT E_EMAIL FROM EMAILS"
	set re = Server.CreateObject("ADODB.Recordset")
	re.ActiveConnection = conn_
	re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 3 : re.Open()

		numBuenos = 0
		arr = split(buenos,"|")
		for each email in arr
			if validarEmail(email) = true then
				re.addnew()
				re("E_EMAIL") = email
				numBuenos = numBuenos + 1
			end if
		next
	re.update()
	re.Close()
	set re = nothing
	end if

	%>
	</tr>
</table>
<%
else
%>

<form name="form1" method="post" action="index.asp">
  <table width="400"  border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#F4F2E8">
    <tr>
      <td align="left"><b>Escribe la cadena de emails</b></td>
    </tr>
    <tr>
      <td align="center"><textarea name="cadena" cols="74" rows="20" wrap="virtual" class="areacadena" id="cadena"></textarea>      </td>
    </tr>
    <tr>
      <td align="right"><input name="" type="submit" style="height:50px" value="Procesar"></td>
    </tr>
  </table>
</form>
	<%end if%>