<%
	ruta_xml_grupos = "/" & c_s & "datos/grupos.xml"
	
	Set xml_grupos = CreateObject("MSXML.DOMDocument")
	if not xml_grupos.Load(Server.MapPath(ruta_xml_grupos)) then
		unerror = true : msgerror = "No ha sido posible cargar la configuración de los grupos."
	Else
		set nodoGrupos = xml_grupos.selectsingleNode("datos/grupos")
		if not typeOK(nodoGrupos) then
			unerror = true : msgerror = "No se ha encontrado ningún grupo."
		end if
	End If
	
	
	function setGrupo(id_grupo)
		if typeOK(nodoGrupos) then
			set setGrupo = nodoGrupos.selectNodes("//datos/grupos/grupo[@id="& numero(id_grupo) &"]").item(0)
		end if
	end function
%>