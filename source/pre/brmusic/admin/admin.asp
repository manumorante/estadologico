<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="inc_rutinas.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>aSkipper</title>
</head>
<body>
<%
' O_o

	Dim cualid
	Dim idi
	Dim id
	
	cualid = ""& request("cualid")
	idi = ""& request("idi")
	id = numero(request("id"))

	' ¿Que hacemos?
	select case request("ac")
	
	' Aplicación: Partes de carga
	case "info_partes"
		

	' Ninguna opción
	case else
		Response.Write "<b>No se ha recibido ninguna petición.</b>"
	end select

%>
</body>
</html>
