<%
	conn_ = ""
	if not unerror then
		conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("\"& c_s &"datos\"& idioma &"\"& cualid &"\"& cualid &".mdb")
	end if
%>