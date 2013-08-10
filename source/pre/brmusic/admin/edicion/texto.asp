<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<!--#include file="bbcode.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Texto</title>
</head>
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
<body class="bodyAdmin">
<%
if not getPermisoParaRuta("edicion", session("idioma"), session("usuario"),request.QueryString("file")) then
	unerror = true : msgerror = "Usted no tiene derechos de aSkipper para esta zona."
end if

Dim idi
Dim archivo
Dim num
Dim archivoXml

idi = ""&request("idi")
archivo = ""&request("archivo")
num = ""&request("num")

if idi = "" or archivo = "" or num = "" then
	unerror = true : msgerror = "No se han recibido todos los parámetros necesarios."
end if

if not unerror then
	archivoXml = "../../" & idi & archivo & "/" &nombrearchivo(archivo) & ".xml"
	dim xmlObj
	set xmlObj = CreateObject("MSXML.DOMDocument")
	if not xmlObj.Load(Server.MapPath(archivoXml)) then
		unerror = true : msgerror = "No se ha encontrado el archivo que desea editar."
	end if
end if

if not unerror then
	dim nodoTexto
	set nodoTexto = xmlObj.selectSingleNode("contenido/texto"&num)
	if not typeOK(nodoTexto) then
		unerror = true : msgerror = "No se ha encontrado 'texto"&num&"' en el XML."
	end if	
end if

select case request.QueryString("ac")
case "editar"
	nodoTexto.text = filtroHtml(request("texto"&num))
	xmlObj.save Server.MapPath(archivoXml)
	%>
	<script language="javascript" type="text/javascript">
		parent.opener.location.href=parent.opener.location
		window.close()
	</script>
	<%
'	Response.Redirect("texto.asp?idi="&idi&"&archivo="&archivo&"&num="&num)
case else

	if not unerror then%>
		<form name="form1" action="texto.asp?ac=editar" method="post">
			<input type="hidden" name="idi" value="<%=idi%>">
			<input type="hidden" name="num" value="<%=num%>">
			<input type="hidden" name="archivo" value="<%=archivo%>">
			<table width="100%"  border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td><%=textAreaBbcode ("form1",""&nodoTexto.nodeName,""&nodoTexto.text,"no se usa",27,""&nodoTexto.getAttribute("comentario"))%></td>
			  </tr>
			  <tr>
				<td align="right" valign="top"><input name="" type="button" class="botonAdmin" onClick="window.close()" value="Cerrar">
				<input type="submit" class="botonAdmin" value="Enviar">
				&nbsp;</td>
			  </tr>
			</table>
		</form>
	<%end if

end select


if unerror then
	Response.Write "<b>Error:</b><br>"&msgerror
end if


%>
</body>
</html>
<%
Function nombrearchivo(valor)
	for n=0 to len(valor)-1
		if Mid(valor,len(valor)-n,1)="\" or Mid(valor,len(valor)-n,1)="/" then
			n=len(valor)+1
		else
			nombrearchivo=Mid(valor,len(valor)-n,1)&nombrearchivo
		end if
	next
end Function
%>