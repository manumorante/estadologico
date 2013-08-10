<!--#include file="inc_visorN_Lib.asp" -->
<!--#include virtual="/admin/global/inc_inicia_xml.asp" -->
<%
	sub visor (pCualid, pIdioma)
		Response.Write "<b>"& ucase(pCualid) &"</b> ("& pIdioma &")"
		dim vn ' visor n
		set vn = New VisorN
		vn.cualid = pCualid
		vn.idioma = pIdioma

		vn.activar
		Response.Write vn.tabla
	
		if vn.unerror then
			Response.Write vn.msgerror
		end if
	end sub
%>