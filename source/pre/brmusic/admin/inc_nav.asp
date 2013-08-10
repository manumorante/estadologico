<%

if ""&request("nav") = "movil" then
		nav_sistema = "movil"
	else
		' Detección del sistena con el que se navega.
		dim navegador
		navegador = lcase(""&Request.ServerVariables("HTTP_USER_AGENT"))
		if inStr(navegador,"symbianos") or inStr(navegador,"blackberry") or inStr(navegador,"semc") or inStr(navegador,"vodafone") then
			nav_sistema = "movil"
		else
			nav_sistema = "pc"
		end if
	end if

'	nav_sistema = "movil"	
%>