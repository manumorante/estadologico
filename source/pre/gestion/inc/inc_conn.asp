<%
'on error resume next
public conn_
Set conn_ = Server.CreateObject("ADODB.Connection")
' Internet
' --------------------------------------------------------------------------------------------------
'conn_.open "estadologico"

' Local
' --------------------------------------------------------------------------------------------------
conn_.open "Driver={Microsoft Access Driver (*.mdb)};DBQ="& server.MapPath("\db\estadologico.mdb")



if err<>0 then
	unerror = true : msgerror = "[Conexión] No se ha logrado abrir la base de datos.<br>"& err.description
end if
on error goto 0
%>