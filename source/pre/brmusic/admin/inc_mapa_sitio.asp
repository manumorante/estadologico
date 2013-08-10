<!--#include file="inc_rutinas.asp" -->
<%
	pi = request.ServerVariables("PATH_INFO")
	if inStr(pi,"/esp/index.asp") then
		idioma = "esp"
	elseif inStr(pi,"/eng/index.asp") then
		idioma = "eng"
	elseif inStr(pi,"/fra/index.asp") then
		idioma = "fra"
	elseif inStr(pi,"/deu/index.asp") then
		idioma = "deu"
	elseif inStr(pi,"/ita/index.asp") then
		idioma = "ita"
	else
		if session("idioma") <> "" then
			idioma = session("idioma")
		end if
	end if
	Dim tablaabierta
	tablaabierta = 0

	iniciotabla = "<table width = '100%' border = '0' cellpadding = '1' cellspacing = '4'><tr><td>"
	findetabla = "</td></tr></table>"

function pintaruta(valor,elxml)
	set rutaxml = elxml
	for n = 1 to valor
		Set rutaxml = rutaxml.parentnode
		devuelto = rutaxml.nodename&"/"&devuelto
	next
	pintaruta = devuelto
end function

Sub ver(Nodos,seccion)
	Dim oNodo
	For Each oNodo In Nodos
		nombrenodo = ""
		Set nombretemp = oNodo
		for nsepar = 1 to separador+1
			if nombrenodo = "" then
				nombrenodo = nombretemp.nodename
			else
				nombrenodo = nombretemp.nodename&"/"&nombrenodo
			end if
			Set nombretemp = nombretemp.parentnode
		next
		Set nombretemp = Nothing
		nombrenodoe = oNodo.nodeName
		titulonodo = oNodo.getattribute("titulo")
		permitido = 1
		for ex = 0 to ubound(excluidos)
			if strcomp(ucase(nombrenodoe),ucase(excluidos(ex))) = 0 then		
				permitido = 0
			end if
		next
		
		if permitido and separador<3 then
			if separador = 0 then
				if tabaabierta = 1 then
					response.Write(findetabla) 
					response.Write(iniciotabla)
				else
					response.Write(iniciotabla)
					tabaabierta = 1
				end if
			end if
			
			if (oNodo.getattribute("activo") = "0" or ""&oNodo.getattribute("default")<>"" ) then %>
				<table width = "100%" border = "0" cellpadding = "0" class = "mapasitio_seccion<% = separador%>">
					<tr>
					<td><% = Ucase(oNodo.getattribute("titulo"))%></td>
					</tr>
				</table>
			<% elseif ( oNodo.getattribute("activo")<>"0") and ( ""&oNodo.getattribute("ocultarpublico") = "0" or ""&oNodo.getattribute("ocultarpublico") = "") then%>
				<table width = "100%" border = "0" cellpadding = "0" class = "mapasitio_seccion<% = separador%>">
					<tr>
					<td><a href = "/<%=c_s%><% = idioma%>/index.asp?secc = /<% = nombrenodo%>" <%if ejecutalocal = 1 then response.write("target = _blank") end if%>><% = Ucase(titulonodo)%></a></td>
					</tr>
				</table>
			<% end if
			
			' Si el nodo siguiente tiene hijos y el actual (padre) no es oculto (ocultarpublico)
			If oNodo.hasChildNodes and ""&oNodo.getAttribute("ocultarpublico") <> "1" Then
				separador = separador+1
				ver oNodo.childNodes,oNodo.nodename
				separador = separador-1
			End If
		end if
	Next
End Sub
%>

<table width = "100%" border = "0" cellpadding = "0" cellspacing = "0" bordercolor = "#789AEB">
	<tr>
	<td><%
	Dim valorant, separador, xmlObj
	separador = 0
	Set xmlObj = CreateObject("MSXML.DOMDocument")
	if xmlObj.Load(Server.MapPath("secciones.xml")) then
		ver xmlObj.selectsinglenode("/pagina/secciones").childnodes,"principal"
	Else
		response.Write("Ha ocurrido un error.")
	End If
	response.write(findetabla)
	%>
	</td>
	</tr>
</table>

