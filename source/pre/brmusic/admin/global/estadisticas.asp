<!--#include virtual="/datos/inc_config_gen.asp" -->
<%	Dim cualid : cualid = "bannerportada"%>
<!--#include virtual="/admin/visores/inc_conn.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Documento sin título</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="estilos.css" rel="stylesheet" type="text/css">
</head>

<body>
<img src="<%=request("imagen")%>" >

<br>
<br> 
<strong>DATOS GLOBALES</strong><br>

<% dia1=request("dia1")
dia2=request("dia2")
mes1=request("mes1")
mes2=request("mes2")
annio1=request("annio1")
annio2=request("annio2")%>

<%

if dia1="" or mes1="" or annio1="" then

fecha1 = date
else
fecha1=RIGHT("0"&dia1,2)&"/"&RIGHT("0"&mes1,2)&"/"&annio1
end if


if dia2="" or mes2="" or annio2="" then

fecha2 = date
else
fecha2=RIGHT("0"&dia2,2)&"/"&RIGHT("0"&mes2,2)&"/"&annio2
end if


		set rex = Server.CreateObject("ADODB.Recordset")
		rex.ActiveConnection = conn_
		sql="SELECT DISTINCT ORIGEN FROM ACCESOS WHERE R_ID="&request.QueryString("id")
		rex.Source = sql : rex.CursorType = 3 : rex.CursorLocation = 2 : rex.LockType = 1
		rex.Open()
		totalpulsados=0
		while not rex.eof

	totalvisitasdiferentes=totalvisitasdiferentes+1

		rex.movenext
		wend

		rex.Close()
		sql="SELECT ORIGEN FROM ACCESOS WHERE R_ID="&request.QueryString("id")
		rex.Source = sql : rex.CursorType = 3 : rex.CursorLocation = 2 : rex.LockType = 1
		rex.Open()
		totalpulsados=0
		while not rex.eof

	totalpulsados=totalpulsados+1

		rex.movenext
		wend

		rex.Close()
		
			if totalpulsados>0 then
	response.Write("<br>El enlace ha sido pulsado en total "&totalpulsados&" veces.<br>")
	response.Write("Ha sido visitado por "&totalvisitasdiferentes&" personas diferentes.<br>")


		end if
		
		
		
		sql= "SELECT DISTINCT ORIGEN FROM ACCESOS WHERE (FECHA BETWEEN #"&fecha1&"# AND #"&fecha2&"# OR FECHA LIKE '"&fecha1&"' OR FECHA LIKE '"&fecha2&"') AND R_ID="&request.QueryString("id")		
		rex.Source = sql : rex.CursorType = 3 : rex.CursorLocation = 2 : rex.LockType = 1
		rex.Open()
totalvisitasdiferentes=0
	while not rex.eof
	
	totalvisitasdiferentes=totalvisitasdiferentes+1

	
	rex.movenext
	wend
	rex.close
	sql= "SELECT ORIGEN FROM ACCESOS WHERE (FECHA BETWEEN #"&fecha1&"# AND #"&fecha2&"# OR FECHA LIKE '"&fecha1&"' OR FECHA LIKE '"&fecha2&"') AND R_ID="&request.QueryString("id")		
		rex.Source = sql : rex.CursorType = 3 : rex.CursorLocation = 2 : rex.LockType = 1
		rex.Open()
totalpulsados=0
	while not rex.eof
	
	totalpulsados=totalpulsados+1

	
	rex.movenext
	wend
	%>
<br>
<br>
<form action="#" method="post" name="fechas">
 Entre <nobr><select name="dia1" class="campo">
  <% 
  if dia1<>"" then
  for n=1 to 31 %>
    <option value="<%=n%>" <%if cInt(dia1)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
	<% else
	  for n=1 to 31 %>
    <option value="<%=n%>" <%if day(date)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
<%	end if%>
  </select>
  <select name="mes1" class="campo">
  <%if mes1<>"" then
  for n=1 to 12 %>
    <option value="<%=n%>" <%if cInt(mes1)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
	<% else
	  for n=1 to 12 %>
    <option value="<%=n%>" <%if month(date)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
<%	end if%>
  </select></nobr>
  <select name="annio1" class="campo">
   <%if annio1<>"" then
    for n= 2004 to year(date) %>
    <option value="<%=n%>" <%if cInt(annio1)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
	<% else
	   for n= 2004 to year(date) %>
    <option value="<%=n%>" <%if year(date)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
<%	end if%> 

  </select>
  y
  

   <select name="dia2" class="campo">
  <% 
  if dia2<>"" then
  for n=1 to 31 %>
    <option value="<%=n%>" <%if cInt(dia2)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
	<% else
	  for n=1 to 31 %>
    <option value="<%=n%>" <%if day(date)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
<%	end if%>
  </select>
  <select name="mes2" class="campo">
  <%if mes2<>"" then
  for n=1 to 12 %>
    <option value="<%=n%>" <%if cInt(mes2)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
	<% else
	  for n=1 to 12 %>
    <option value="<%=n%>" <%if month(date)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
<%	end if%>
  </select>
  <select name="annio2" class="campo">
   <%if annio2<>"" then
    for n= 2004 to year(date) %>
    <option value="<%=n%>" <%if cInt(annio2)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
	<% else
	   for n= 2004 to year(date) %>
    <option value="<%=n%>" <%if year(date)=n then response.Write(" selected") end if%>><%=n%></option>
	<%next%>
<%	end if%> 

  </select>
  <input type="submit" name="Submit" value="Enviar" class="boton">  
</form>

Resultados para el rango: <%=fecha1%> y <%=fecha2%><br><br>


<%
	if totalpulsados>0 then
	response.Write("El enlace ha sido pulsado "&totalpulsados&" veces.<br>")
	response.Write("Ha sido visitado por "&totalvisitasdiferentes&" personas diferentes.<br>")

	if concurrentes>=1 then
	response.Write("De las cuales "&concurrentes&" han reincidido.")
	end if
	

	else
	
	response.Write("El anuncio no ha sido visitado en el rango de fechas especificado.")

	end if	
rex.close
set rex=nothing

%>
</body>
</html>
