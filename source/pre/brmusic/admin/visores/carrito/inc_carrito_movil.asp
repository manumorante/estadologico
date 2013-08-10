<%



Dim unerror, msgerror
unerror = false : msgerror = ""

secc = ""& request.QueryString("secc")

%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="admin/usuarios/rutinasParaAdmin.asp" -->
<!--#include virtual="/admin/global/inc_rutinas.asp" -->
<!--#include file="inc_xmlShop.asp" -->
<!--#include virtual="/admin/inc_sendmail.asp" -->

<%

set carrito = New xmlSessionShop

dim idioma
dim cualid
dim id
dim ruta_conn_productos
dim conn_activa
dim re, reTotal

' Tomo los datos minimos necesários para localizar el artículo/servicio en la base de datos (MDB).
idioma = ""& request("idi")
cualid = ""&request("cualid")
id = numero(request("id"))

ruta_conn_productos = ""
if idioma<>"" and cualid<>"" and id<>0 then
	ruta_conn_productos = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("/"& c_s &"datos/"& idioma &"/"& cualid &"/"& cualid &".mdb")
end if

' Genero la conexión a la MDB de productos
if ruta_conn_productos <> "" then
	set conn_activa = Server.CreateObject("ADODB.Connection")
	on error resume next
		conn_activa.open ruta_conn_productos
		if err=0 then conn_aciva_ok = true else conn_aciva_ok = false end if
	on error goto 0
end if


select case request.QueryString("ac")
case "setopt"
	cantidad = numero(request.QueryString("cantidad"))
	if cantidad>0 then
		call carrito.addAtt(idioma, cualid, id, "cantidad", cantidad)
		call carrito.addAtt(idioma, cualid, id, "comentarios", ""& request.QueryString("comentarios"))
		Response.Redirect("index.asp?secc=/carrito&seccrefer="& request.QueryString("seccrefer") &"&idsecc="& request.QueryString("idsecc") &"&idsecc2="& request.QueryString("idsecc2") &"")
	end if
case "opt"
	if not unerror then
		sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
		consultaXOpen sql,1
		if reTotal > 0 then
			titulo = ""& re("R_TITULO")
			precio = re("R_PRECIO")
		end if
		consultaXClose()
	%>
<h2><%=titulo%></b></h2>
	<form action="index.asp" method="get" name="f" id="f">
	<input type="hidden" name="secc" value="/carrito" />
	<input type="hidden" name="ac" value="add" />
	<input type="hidden" name="idi" value="<%=idioma%>" />
	<input type="hidden" name="cualid" value="<%=cualid%>" />
	<input type="hidden" name="id" value="<%=id%>" />
	<input type="hidden" name="seccrefer" value="<%=request.QueryString("seccrefer")%>" />
	<input type="hidden" name="idsecc" value="<%=request.QueryString("idsecc")%>" />
	<input type="hidden" name="idsecc2" value="<%=request.QueryString("idsecc2")%>" />
	<input type="hidden" name="pag" value="<%=request.QueryString("pag")%>" />
	
	<%if ""& request.QueryString("msg") <> "" then
		Response.Write "<em>"& request.QueryString("msg") &"</em><br />"
	end if%>
	Cantidad: <br />
	<input type="text" value="1" name="cantidad" /><br />
	Comentarios: 
	<br />
	<textarea name="comentarios"></textarea>
	<br />
	<div align="right">
	  <input type="submit" value="Aceptar" />
	  <br />
	  <a href="index.asp?secc=<%=request.QueryString("seccrefer")%>&ac=listado&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>&pag=<%=request.QueryString("pag")%>">Cancelar</a></div>
</form>
<%
	end if
	
case "add"
	if not unerror then
		sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
		consultaXOpen sql,1
		if reTotal > 0 then
			titulo = ""& re("R_TITULO")
			precio = re("R_PRECIO")
		end if
		consultaXClose()
	end if
	
	cantidad = numero(request.QueryString("cantidad"))
	if cantidad=0 then
		Response.Redirect("index.asp?secc=/carrito&idi="& idioma &"&cualid="& cualid &"&id="& id &"&seccrefer="& request.QueryString("seccrefer") &"&ac=opt&idsecc="& request.QueryString("idsecc") &"&idsecc2="& request.QueryString("idsecc2") &"&msg=Escriba una cantidad correcta.")
	end if
	call carrito.addItem(idioma, cualid, id, titulo, precio, cantidad)
	call carrito.addAtt(idioma, cualid, id, "comentarios", ""&request.QueryString("comentarios"))
	
	if carrito.unerror then
		Response.Write "<br />"& carrito.msgerror
	else
		Response.Redirect("index.asp?secc="& request.QueryString("seccrefer") &"&ac=listado&idsecc="& request.QueryString("idsecc") &"&idsecc2="& request.QueryString("idsecc2")) &"&pag="& request.QueryString("pag")
	end if

case "eliminar"
	call carrito.delItem(idioma, cualid, id)
	Response.Redirect("index.asp?secc=/carrito&seccrefer="& request.QueryString("seccrefer") &"&idsecc="& request.QueryString("idsecc") &"&idsecc2="& request.QueryString("idsecc2"))
case "info"

	if not conn_aciva_ok then
		Response.Write "<b>No se ha encontrado el artículo.</b>"
	else
		sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id
		consultaXOpen sql,1
		if reTotal >= 1 then
			cantidad = carrito.getAtt(idioma, cualid, id,"cantidad")
		%>
		<div align="left"><h2><%=re("R_TITULO")%></h2></div>
		<div align="right"><%=euros(re("R_PRECIO"))%> &euro;</div>
		<form name="f_cantidad" method="get" action="index.asp">
		<input type="hidden" name="secc" value="/carrito" />
		<input type="hidden" name="ac" value="setopt" />
		<input type="hidden" name="idi" value="<%=idioma%>" />
		<input type="hidden" name="cualid" value="<%=cualid%>" />
		<input type="hidden" name="id" value="<%=id%>" />
		<input type="hidden" name="seccrefer" value="<%=request.QueryString("seccrefer")%>" />
		<input type="hidden" name="pag" value="<%=request.QueryString("pag")%>" />
		Cantidad:<br />
		<input name="cantidad" type="text" class="campo" id="cantidad" value="<%=cantidad%>">
		<br />
		Comentarios: 
	    <br />
        <textarea name="comentarios"><%=carrito.getAtt(idioma, cualid, id,"comentarios")%></textarea>
        <br />
        <input type="submit" value="Aceptar" />
		</form>
		<br />
		<a href="index.asp?secc=/carrito&ac=eliminar&idi=<%=idioma%>&cualid=<%=cualid%>&id=<%=id%>&seccrefer=<%=request.QueryString("seccrefer")%>">[X] Eliminar del carrito</a><br />
		<%end if
		consultaXClose()
	end if

case "confirm"

	cif = ""& request.Form("cif")
	emailcopia = ""&request.Form("emailcopia")
	
	if cif = "" then
		Response.Redirect("index.asp?secc=/carrito&ac=datped&seccrefer="& request.Form("seccrefer") &"&idsecc="& request.Form("idsecc") &"&idsecc2="& request.Form("idsecc2") &"&msg=Por favor, rellene el campo CIF.")
	end if

	' Leer variables de configuración.
	'---------------------------------------------------------------------------------------------
	ruta_xml_admindatos = "/"& c_s &"datos/esp/admindatos_carrito/admindatos_carrito.xml"
	set xml_admindatos = CreateObject("MSXML.DOMDocument")
	if not xml_admindatos.Load(Server.MapPath(ruta_xml_admindatos)) then
		unerror = true : msgerror = "'XML admindatos' Error de carga."
	else
		set nodo_admindatos = xml_admindatos.selectSingleNode("datos")
		if not typeOK(nodo_admindatos) then
			unerror = true : msgerror = "'No se ha configurado la tienda online."
		end if
	end if
	
	Dim contador
	Dim numerodeceros
	Dim prefijopedido

	' Contador
	if not unerror then
		set nodo_contador = nodo_admindatos.selectSingleNode("contadordereferencias")
		if not typeOK(nodo_contador) then
			unerror = true : msgerror = "No está configurado el nodo contador de referencias."
		else
			contador = numero(nodo_contador.text)
			' Incrementar el contador
			nodo_contador.text = numero(nodo_contador.text)+1
		end if
	end if

	' Número de ceros
	if not unerror then
		set nodo_numerodeceros = nodo_admindatos.selectSingleNode("numerodeceros")
		if not typeOK(nodo_numerodeceros) then
			unerror = true : msgerror = "No está configurado la cantidad de ceros para las referencias de los pedidos."
		else
			numerodeceros = numero(nodo_numerodeceros.text)
		end if
	end if

	' Prefijo para la referencia
	if not unerror then
		set nodo_prefijopedido = nodo_admindatos.selectSingleNode("prefijopedido")
		if not typeOK(nodo_prefijopedido) then
			unerror = true : msgerror = "No está configurado un prefijo para las referencias de los pedidos."
		else
			prefijopedido = ""& nodo_prefijopedido.text
		end if
	end if

	' Nombre Email de emisión
	if not unerror then
		set nodo_nombreemailemision = nodo_admindatos.selectSingleNode("nombreemailemision")
		if not typeOK(nodo_nombreemailemision) then
			unerror = true : msgerror = "No está configurado un nombre del email de emisión."
		else
			nombreemailemision = ""& nodo_nombreemailemision.text
		end if
	end if

	' Email de emisión
	if not unerror then
		set nodo_emailemision = nodo_admindatos.selectSingleNode("emailemision")
		if not typeOK(nodo_emailemision) then
			unerror = true : msgerror = "No está configurado el email de emisión."
		else
			emailemision = ""& nodo_emailemision.text
		end if
	end if

	' Nombre Email de recepción
	if not unerror then
		set nodo_nombreemailrecepcion = nodo_admindatos.selectSingleNode("nombreemailrecepcion")
		if not typeOK(nodo_nombreemailrecepcion) then
			unerror = true : msgerror = "No está configurado el nombre del email de recepción."
		else
			nombreemailrecepcion = ""& nodo_nombreemailrecepcion.text
		end if
	end if

	' Email de recepción
	if not unerror then
		set nodo_emailrecepcion = nodo_admindatos.selectSingleNode("emailrecepcion")
		if not typeOK(nodo_emailrecepcion) then
			unerror = true : msgerror = "No está configurado el email de recepción."
		else
			emailrecepcion = ""& nodo_emailrecepcion.text
		end if
	end if
	'---------------------------------------------------------------------------------------------
	
	if not unerror then
		' Guardo el xml de configuración
		xml_admindatos.Save Server.MapPath(ruta_xml_admindatos)
	
		set nodo_contador = nothing
		set nodo_admindatos = nothing
		set nodo_numerodeceros = nothing
		set xml_admindatos = nothing
		
		' Generar númeo de pedido.
		referencia = prefijopedido & right("00000000000" & (contador+1),numerodeceros)
	end if
	
	' General email
	'---------------------------------------------------------------------------------------------
	if not unerror then
		nombre_usuario = ""&getNombreUsuario(session("usuario"))
	
		Dim str_pedido
		str_pedido = "<table width=550 border=0 cellspacing=0 cellpadding=0>"
		str_pedido = str_pedido & "<tr><td colspan=4><h3>Información sobre su pedido nº"& referencia &"</h3></td></tr>"
		str_pedido = str_pedido & "<tr><td colspan=4><b>Pedido realizado por:</b> "& nombre_usuario &"</td></tr>"
		str_pedido = str_pedido & "<tr><td colspan=4><b>Cliente:</b> "& cif &"</td></tr>"
		str_pedido = str_pedido & "<tr><td colspan=4><b>Fecha:</b> "& Date() &"</td></tr>"
		str_pedido = str_pedido & "<tr><td colspan=4>&nbsp;</td></tr>"
		total = 0
		str_pedido = str_pedido & "<tr><td><b>Ref.</b></td><td><b>Art&iacute;culo</b></td><td><b>Cantidad</b></td><td align=right><b>Precio</b></td></tr>"
		str_pedido = str_pedido & "<tr><td colspan=4><hr></td></tr>"
		for each item in carrito.nodo.childNodes
			titulo = ""& item.getAttribute("titulo")
			cantidad = numero(item.getAttribute("cantidad"))
			comentarios = ""& item.getAttribute("comentarios")
			precio = numero(item.getAttribute("precio"))
			ref = ""& item.getAttribute("referencia")
			subtotal = precio*cantidad
			total = total + subtotal
	
			str_pedido = str_pedido & "<tr>"
			str_pedido = str_pedido & "<td>"& ref &"</td>"
			str_pedido = str_pedido & "<td>"& titulo
			if comentarios <> "" then
				str_pedido = str_pedido & " ("& comentarios &")"
			end if
			str_pedido = str_pedido & "</td>"
			str_pedido = str_pedido & "<td>"& cantidad &"</td>"
			str_pedido = str_pedido & "<td align=right>"& subtotal & " &euro;</td></tr>"
		next
		str_pedido = str_pedido & "<tr><td colspan=4><hr></td></tr>"
		str_pedido = str_pedido & "<tr><td colspan=4 align=right>Subtotal: "& euros(total) & " &euro;</td></tr>"
		iva = (total * 16)/100
		str_pedido = str_pedido & "<tr><td colspan=4 align=right>IVA: "& euros(iva) & " &euro;</td></tr>"
		str_pedido = str_pedido & "<tr><td colspan=4 align=right><b>Total</b>: "& euros(total+iva) & " &euro;</td></tr>"

		str_pedido = str_pedido & "</table>"
	end if
'	 Response.Write str_pedido
	'---------------------------------------------------------------------------------------------

	' Enviar email a la central.
	'---------------------------------------------------------------------------------------------
	Subject = "Su pedido no.: " & referencia & " - Por: "& nombre_usuario &"."
	Body = str_pedido
	if sendMail(emailemision, nombreemailemision, emailrecepcion, nombreemailrecepcion, Subject, Body) then
'		Response.Write "ok"
	else
'		Response.Write "ko"
	end if
	
	' Enviar una copia al email indicado
	if emailcopia <> "" then
		call sendMail(emailemision, nombreemailemision, emailcopia, emailcopia, Subject, Body)
	end if
	'---------------------------------------------------------------------------------------------

	' Enviar copia al cliente, si se solicita.

	' Insertar el pedido
	'---------------------------------------------------------------------------------------------
	if not unerror then
		ruta_conn_pedidos = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("/"& c_s &"datos/esp/pedidos/pedidos.mdb")
		set conn_activa = Server.CreateObject("ADODB.Connection")
		on error resume next
			conn_activa.open ruta_conn_pedidos
		on error goto 0
	
		titulo = referencia &" - "& cif
		seccion = 1
		seccion2 = 1
		usuario = numero(session("usuario"))	' ID del usuario logeado (ej: comercial que ordena el pedido)
		fuente = ""
		alfinal = 1
		enportada = 0
		activo = 1
		enlace = ""
		fecha = Date()					' Fecha en la que realiza el pedido
		fechaini = ""
		fechafin = ""
		resto_nombres = ", R_MEMO1, R_REF"
		resto_valores = ",'"& str_pedido &"', '"& referencia &"'"
		conn = ruta_conn_pedidos
	
		str_insertar = ""& InsertarRegistro (titulo, seccion, seccion2, usuario, fuente, alfinal, enportada, activo, enlace, fecha, fechaini, fechafin, resto_nombres, resto_valores, conn)
	end if
	'---------------------------------------------------------------------------------------------


	%>
<div align="left">
<h1>Pedido confirmado</h1>
<p>Su pedido ha sido tramitado con &eacute;xito.</p>
<p><a href="index.asp?secc=<%=request.Form("seccrefer")%>">Men&uacute; principal</a><br />
  <a href="index.asp?secc=<%=request.Form("seccrefer")%>&ac=listado">Listado completo</a></p>
</div>
	  <%case "datped"%>
    
      <b>Datos del ciente</b><br />

	<form action="index.asp?secc=/carrito&ac=confirm" method="post" name="f">
<input type="hidden" name="idi" value="<%=idioma%>"/>
<input type="hidden" name="cualid" value="<%=cualid%>"/>
<input type="hidden" name="seccrefer" value="<%=request.QueryString("seccrefer")%>"/>
<input type="hidden" name="idsecc" value="<%=request.QueryString("idsecc")%>"/>
<input type="hidden" name="idsecc2" value="<%=request.QueryString("idsecc2")%>"/>
<p>
<%if ""& request.QueryString("msg") <> "" then%>
	<p><div align="center"><em><%=request.QueryString("msg")%></em></div></p>
<%end if%>
CIF: 
<br />
<input name="cif" type="text" id="cif">
<br />
Enviar copia a:<br />
<input name="emailcopia" type="text" id="emailcopia">
</p>

	<%if carrito.nodo.childNodes.length <= 0 then%>
		<b>Ning&uacute;n art&iacute;culo en carrito.</b>
	<%else
		for each item in carrito.nodo.childNodes
			cantidad = numero(item.getAttribute("cantidad"))
			cantidad = numero(item.getAttribute("cantidad"))
			precio = numero(item.getAttribute("precio"))
			subtotalA = precio * cantidad
			subtotal = subtotal + subtotalA
		next
		
		iva = (16*subtotal)/100
		total = subtotal + iva
		%>
		<hr noshade size="1"/>
		<div align="right">
		Subtotal: <%=euros(subtotal)%> &euro;<br />
		IVA: <%=euros(iva)%> &euro;<br />
		<b>Total:</b> <%=euros(total)%> &euro;
		<hr noshade size="1"/>

		<p><input type="submit" value="Aceptar" /></p>
		</div>
	<%end if%>
</form>
<%case else%>


	<%if carrito.nodo.childNodes.length <= 0 then%>
		<b>Ning&uacute;n art&iacute;culo en carrito.</b>
	<%else%>
		<hr noshade size="1" />
		<%for each item in carrito.nodo.childNodes
			cantidad = numero(item.getAttribute("cantidad"))
			cantidad = numero(item.getAttribute("cantidad"))
			%>
			<%=cantidad%> x <span title="ID:<%=item.getAttribute("id")%>"><a href="index.asp?secc=<%=secc%>&ac=info&idi=<%=item.getAttribute("idioma")%>&cualid=<%=item.getAttribute("cualidad")%>&id=<%=item.getAttribute("id")%>&seccrefer=<%=request.QueryString("seccrefer")%>"><%=item.getAttribute("titulo")%></a></span>
			<%
			precio = numero(item.getAttribute("precio"))
			
			subtotalA = precio * cantidad
			subtotal = subtotal + subtotalA
			Response.Write euros(subtotalA)%> &euro;
			<br />
		<%next
		
		iva = (16*subtotal)/100
		total = subtotal + iva
		%>
		<hr noshade size="1" />
		<div align="right"><br />
		Subtotal: <%=euros(subtotal)%> &euro;<br />
		IVA: <%=euros(iva)%> &euro;<br />
		<b>Total:</b> <%=euros(total)%> &euro;
		</div>

	<%end if%>
	<br /><br />
	<%if ""& request.QueryString("seccrefer") <> "" then%>
		<a href="index.asp?secc=<%=request.QueryString("seccrefer")%>&ac=listado&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>">&lt;&lt; Seguir comprando</a><br />
	<%else%>
		<a href="<%=request.ServerVariables("HTTP_REFERER")%>">&lt;&lt; Seguir comprando</a><br />
	<%end if%>
	<a href="index.asp?secc=/carrito&ac=datped&seccrefer=<%=request.QueryString("seccrefer")%>&idsecc=<%=request.QueryString("idsecc")%>&idsecc2=<%=request.QueryString("idsecc2")%>">Confirmar pedido &gt;&gt;</a>


    <%end select

if unerror then
	Response.Write "<b>Error:</b><br />"& msgerror
end if


		
%>

	