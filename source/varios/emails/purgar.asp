<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="inc_lib_email.asp" -->
<%
	Dim conn_
	conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("../../data/contactos.mdb")
%>
<html>
<head>
<title>Purgar</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 7.5pt;
}
body {
	background-color: #ECE9D8;
}
-->
</style></head>
<body>
<%
	' Email devueltos
	contenido = ""
	Set Upload = Server.CreateObject("Persits.Upload")
	Set Dir = Upload.Directory(Server.MapPath(".") & "/emailsmalos/*.*")
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	numEmailsMalos = 0
	For Each item in Dir
		nombre = lcase(item.filename)
		nombre = sinNum(nombre)
		nombre = replace(nombre,"(","")
		nombre = ""&replace(nombre,")","")

'		if inStr(nombre,"failure") or inStr(nombre,"borrar") or inStr(nombre,"fallo en la entrega") or inStr(nombre,"returned")  or inStr(nombre,"not send") or inStr(nombre,"returning") or inStr(nombre,"undeliverable") or inStr(nombre,"deliver") or inStr(nombre,"Inexistente") or inStr(nombre,"not exist") or inStr(nombre,"unknown") then
		if nombre <> "" then
			archivo = Server.MapPath("emailsmalos/"&item.Filename)
			Response.Write archivo & "<br>"
			set leer = fso.OpenTextFile(archivo,2,false)
			contenido = contenido & lcase(leer.Readall)
			numEmailsMalos = numEmailsMalos + 1
			set leer = nothing
'			call borrarArchivo(archivo)
		end if
	Next
	Set Dir = nothing
	malos = getEmails(contenido)
	
	Response.Write "<br>Devueltos o asunto 'BORRAR': "& numEmailsMalos
	
	numMalosEnBuenos = 0
	if numEmailsMalos > 0 then
		' Insertar los email malos en la tabla de malos
		sql = "SELECT EM_EMAIL FROM MALOS"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 3 : re.Open()
	
			arr = split(malos,"|")
			for each a in arr
				re.addnew()
				re("EM_EMAIL") = a
			next
		re.update()
		re.Close()
		set re = nothing
		
		' Borrar esos email de la lista buena (por si estubiesen ...)
		sql = "SELECT E_EMAIL FROM EMAILS"

		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 3 : re.Open()

		while not re.eof 
			if inStr(malos,re("E_EMAIL")) then
				re.delete()
				numMalosEnBuenos = numMalosEnBuenos + 1
			end if
			re.movenext
		wend

		re.Close()
		set re = nothing
		
	end if

	Response.Write "<br>Malos borrados de la tabla buenos: "& numMalosEnBuenos

	' Hago una limpieza extra de los emails malos en la lista buena
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
	
	sql = "SELECT E_EMAIL FROM EMAILS"
	
	set re = Server.CreateObject("ADODB.Recordset")
	re.ActiveConnection = conn_
	re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 3 : re.Open()

	borrados = 0
	while not re.eof 
		if inStr(malos,re("E_EMAIL")) then
			re.delete()
			borrados = borrados + 1
		end if
		re.movenext
	wend

	re.Close()
	set re = nothing
	
	Response.Write "<br>Borrados de lista buena de lista mala: " & borrados
	
	
	' Cargo una lista con todos los email
	sql = "SELECT DISTINCT E_EMAIL FROM EMAILS"
	
	set re = Server.CreateObject("ADODB.Recordset")
	re.ActiveConnection = conn_
	re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 3 : re.Open()

	todos = "|"
	while not re.eof 
		todos = todos & re("E_EMAIL") &"|"
		re.movenext
	wend

	re.Close()
	set re = nothing
	
	Response.Write "<br>Distintos: ("& cuentaPalabras(todos,"@") &")"
	
	' Leo de nuevo la bd y borros los que aparezcan dos o mas veces en la lista
	sql = "SELECT E_EMAIL FROM EMAILS"
	
	set re = Server.CreateObject("ADODB.Recordset")
	re.ActiveConnection = conn_
	re.Source = sql : re.CursorType = 1 : re.CursorLocation = 1 : re.LockType = 3 : re.Open()

	lista1 = ""
	numRepetidos = 0
	while not re.eof 
		if not validarEmail(re("E_EMAIL")) = true then
			re.delete
		else
			if inStr(lista1,re("E_EMAIL")) then
				re.delete
				numRepetidos = numRepetidos + 1
			else
				lista1 = lista1 & re("E_EMAIL")
			end if
		end if
		re.movenext
	wend

	re.Close()
	set re = nothing
	
	Response.Write "<br>Repetidos o incorrectos borrados: ("& numRepetidos &")"
	
	
	
	
	
%>
</body>
</html>
