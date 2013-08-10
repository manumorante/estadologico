<!--#include file="inc_sha256.asp" -->
<%
		if ""&request.Form() <> "" then
			c_usuario = request.Form("usuario")
			c_clave = SHA256(request.Form("clave"))
			if c_usuario <> "" and c_clave <> "" then
				if c_clave = getClave(c_usuario) then
					' Declaro la sesion de usuario con su código
					session("usuario") = getCodigo(c_usuario)
					session("idioma") = "esp" ' Español por defecto.
		
					if ""&request.Form("secc_destino") <> "" then
						Response.Redirect("index.asp?secc=" & ""&request.Form("secc_destino"))
					else
						Response.Redirect("index.asp?secc=/inicio")
					end if
					
				else
					errorvalidar = true : msgerrorvalidar = "Usuario o clave incorrectos."
				end if
			else
				errorvalidar = true : msgerrorvalidar = "Debe escribir su nombre y clave de usuario."
			end if
		end if
%>