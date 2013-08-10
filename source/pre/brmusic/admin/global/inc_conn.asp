
<%
	' *** DECLARACIÓN DE LA CONEXION
	public conn_
	public conn_activa
	public config_idioma_bd

	if ""&session("cualid") <> "" and ""&session("idioma") <> "" then
		config_idioma_bd = ""&eval("config_idioma_"& session("idioma") &"_bd")
		if config_idioma_bd = "" then
			carpeta_idioma = session("idioma")
		else
			carpeta_idioma = config_idioma_bd
		end if
		conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("/" & c_s & "datos/"& carpeta_idioma &"/"& session("cualid") &"/"&session("cualid")&".mdb")
		if config_idioma_bd = "esp" and config_idioma_bd = session("idioma") then
			config_idioma_bd = ""
		end if			
		set conn_activa = server.CreateObject("ADODB.Connection")
		on error resume next
		conn_activa.open conn_
		if err<>0 then
			unerror = true : msgerror = "No se ha encontrado la base de datos.<br><b>CONN:</b> " & conn_
		end if
		on error goto 0		
	else
		conn_ = ""
	end if
%>