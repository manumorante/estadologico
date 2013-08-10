<%
on error resume next
	set conn_ = server.CreateObject("ADODB.Connection")
	conn_.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("\"& c_s &"datos\"& idioma &"\"& cualid &"\"& cualid &".mdb")
	if err<>0 then
		unerror = true : msgerror = "No se ha logrado conectar con la base de datos."
	end if
on error goto 0
%>