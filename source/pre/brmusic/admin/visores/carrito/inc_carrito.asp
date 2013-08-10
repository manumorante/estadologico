<%



Dim unerror, msgerror
unerror = false : msgerror = ""

secc = ""& request.QueryString("secc")

%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
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
		Response.Redirect("index.asp?secc=/carrito&seccrefer="& request.QueryString("seccrefer") &"&seccion="& request.QueryString("seccion") &"&seccion2="& request.QueryString("seccion2") &"")
	end if
case "opt"
	if not unerror then
		sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
		consultaXOpen sql,1
		if reTotal > 0 then
			titulo = ""& re("R_TITULO")
			precio = re("R_PVP")
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
	<input type="hidden" name="seccion" value="<%=request.QueryString("seccion")%>" />
	<input type="hidden" name="seccion2" value="<%=request.QueryString("seccion2")%>" />
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
	  <a href="index.asp?secc=<%=request.QueryString("seccrefer")%>&ac=listado&seccion=<%=request.QueryString("seccion")%>&seccion2=<%=request.QueryString("seccion2")%>&pag=<%=request.QueryString("pag")%>">Cancelar</a></div>
</form>
<%
	end if
	
case "add"
	if not unerror then
		sql = "SELECT * FROM REGISTROS WHERE R_ID = "& id &""
		consultaXOpen sql,1
		if reTotal > 0 then
			titulo = ""& re("R_TITULO")
			precio = re("R_PVP")
		end if
		consultaXClose()
	end if
	
	cantidad = numero(request.QueryString("cantidad"))
	if cantidad=0 then
		Response.Redirect("index.asp?secc=/carrito&idi="& idioma &"&cualid="& cualid &"&id="& id &"&seccrefer="& request.QueryString("seccrefer") &"&ac=opt&seccion="& request.QueryString("seccion") &"&seccion2="& request.QueryString("seccion2") &"&msg=Escriba una cantidad correcta.")
	end if
	call carrito.addItem(idioma, cualid, id, titulo, precio, cantidad)
	'[revisar]
'	call carrito.addAtt(idioma, cualid, id, "comentarios", ""&request.QueryString("comentarios"))
	
	if carrito.unerror then
		Response.Write "<br />"& carrito.msgerror
	else
		ir = "index.asp?secc=/carrito&seccrefer="& request.QueryString("seccrefer") &"&seccion="& request.QueryString("seccion") &"&seccion2="& request.QueryString("seccion2")&"&pag="& request.QueryString("pag")
		Response.Redirect(ir)
	end if

case "eliminar"
	call carrito.delItem(idioma, cualid, id)
	Response.Redirect("index.asp?secc=/carrito&seccrefer="& request.QueryString("seccrefer") &"&seccion="& request.QueryString("seccion") &"&seccion2="& request.QueryString("seccion2"))
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
		<div align="right"><%=euros(re("R_PVP"))%> &euro;</div>
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

case "vistafinal"


	' Leer variables de configuración.
	'---------------------------------------------------------------------------------------------
	ruta_xml_admindatos = "/"& c_s &"datos/esp/admindatos_carrito/admindatos_carrito.xml"
	set xml_admindatos = CreateObject("MSXML.DOMDocument")
	if not xml_admindatos.Load(Server.MapPath(ruta_xml_admindatos)) then
		unerror = true : msgerror = "'XML admindatos' Error de carga."
	else
		set nodo_admindatos = xml_admindatos.selectSingleNode("datos")
		if not typeOK(nodo_admindatos) then
			unerror = true : msgerror = "No se ha configurado la tienda online."
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

	' Gastos envio España
	if not unerror then
		set nodo_gastosenvioespana = nodo_admindatos.selectSingleNode("gastosenvioespana")
		if not typeOK(nodo_gastosenvioespana) then
			unerror = true : msgerror = "No está configurado el valor para gastos de envio en España."
		else
			gastosenvioespana = numero(nodo_gastosenvioespana.text)
		end if
	end if

	' Gastos envio Europa
	if not unerror then
		set nodo_gastosenvioeuropa = nodo_admindatos.selectSingleNode("gastosenvioeuropa")
		if not typeOK(nodo_gastosenvioespana) then
			unerror = true : msgerror = "No está configurado el valor para gastos de envio en Europa."
		else
			gastosenvioeuropa = numero(nodo_gastosenvioeuropa.text)
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

	if not unerror then
		nombre_usuario = ""&getNombreUsuario(session("usuario"))
	end if

	str = ""

	str = str & "<table width=100%  border=0 cellspacing=4 cellpadding=10>"
	str = str & "<tr><td><font size=+1><b>Datos de facturación</b></font></td>"
	if ""& request.Form("calle_envio") <> "" then
		str = str & "<td>&nbsp;</td><td><font size=+1><b>Datos de envio</b></font></td>"
	end if
	str = str & "</tr>"

	str = str & "<tr>"
	str = str & "<td bgcolor=#FFFFFF>"
	str = str & request.Form("nombre") &" "& request.Form("apellidos") &"<br>"

	str = str & request.Form("calle") &" "& request.Form("numero") &"<br>"
	str = str & request.Form("cp") &" "& request.Form("ciudad") &"<br>"
	str = str & request.Form("pais") &"<br>"

	str = str & "<br><b>Contacto:</b><br>"
	str = str & request.Form("telefono") &"<br>"
	str = str & request.Form("email") &"<br>"
	str = str & "</td>"

	if ""& request.Form("calle_envio") <> "" then
		str = str & "<td width=5>&nbsp;</td>"

		str = str & "<td bgcolor=#FFFFFF>"
		str = str & request.Form("nombre_envio") &" "& request.Form("apellidos_envio") &"<br>"
	
		str = str & request.Form("calle_envio") &" "& request.Form("numero_envio") &"<br>"
		str = str & request.Form("cp_envio") &" "& request.Form("ciudad_envio") &"<br>"
		str = str & request.Form("pais_envio") &"<br>"
	
		str = str & "<br><b>Contacto:</b><br>"
		str = str & request.Form("telefono_envio") &"<br>"
		str = str & request.Form("email_envio") &"<br>"
		str = str & "</td>"
	end if

	str = str & "</tr>"
	str = str & "</table>"

	str = str & "<br><hr size=1 noshade>"
	str = str & "<table width=100% >"
	str = str & "<tr><td><b>Artículo</b></td><td><b>Cantidad</b></td><td align=right><b>Precio</b></td></tr>"
	
	n=0
	for each item in carrito.nodo.childNodes
		n = n+1
		cantidad = numero(item.getAttribute("cantidad"))
		cantidad = numero(item.getAttribute("cantidad"))
		precio = numero(item.getAttribute("precio"))
		subtotalA = precio * cantidad
		subtotal = subtotal + subtotalA
		
		str = str & "<tr>"
		str = str & "<td>"& item.getAttribute("titulo") &"</td><td>"& cantidad &"</td><td align=right>"& euros(subtotalA) &" &euro;</td>"
		str = str & "</tr>"
	next
	str = str & "</table>"
	str = str & "<hr size=1 noshade>"

	paises_eu = "|AL|AD|AM|AT|BE|BA|BG|HR|CZ|DK|EE|FI|FR|GE|DE|GR|IS|IE|IT|LI|LT|LU|MT|NO|PL|PT|RO|RU|SM|SI|ES|SE|"
	if request.Form("pais") = "ES" then
		gastosenvio = gastosenvioespana
	elseif inStr(paises_eu,"|"& request.Form("pais") &"|") >0 then
		gastosenvio = gastosenvioeuropa
	else
		' No se puede pagar online sin consultar
		gastosenvio = 0
	end if

	str = str &"<div align=right>"
	str = str &"Subtotal: "& euros(subtotal) &" &euro;<br>"
	subtotal = subtotal + gastosenvio
	if gastosenvio > 0 then
		str = str &"Gastos de envio: "& euros(gastosenvio) &" &euro;<br>"
	else
		str = str &"Pendiente gastos de envio<br>"
	end if

	if request.Form("pais") = "ES" then
		iva = (subtotal*16)/100
		str = str &"I.V.A: "& euros(iva) &" &euro;<br>"

	else
		iva = 0
	end if
	total = subtotal+iva
	str = str &"<b>Total</b>: "& euros(total) &" &euro;<br>"
	str = str &"</div>"

	session("str_pedido") = str
	Response.Write str
	
	%>
	<form name="f" action="index.asp?secc=/carrito&ac=confirm" method="post">

		<input type="hidden" name="referencia" value="<%=referencia%>">
		<input type="hidden" name="nombreemailrecepcion" value="<%=nombreemailrecepcion%>">
		<input type="hidden" name="emailrecepcion" value="<%=emailrecepcion%>">
		<input type="hidden" name="gastosenvio" value="<%=gastosenvio%>">
		<input type="hidden" name="total" value="<%=total%>">
	
		<input type="hidden" name="iva" value="<%=iva%>">
		<input type="hidden" name="formapago" value="<%=request.Form("formapago")%>">
		<input type="hidden" name="nombre" value="<%=request.Form("nombre")%>">
		<input type="hidden" name="nif" value="<%=request.Form("nif")%>">
		<input type="hidden" name="apellidos" value="<%=request.Form("apellidos")%>">
		<input type="hidden" name="calle" value="<%=request.Form("calle")%>">
		<input type="hidden" name="numero" value="<%=request.Form("numero")%>">
		<input type="hidden" name="cp" value="<%=request.Form("cp")%>">
		<input type="hidden" name="ciudad" value="<%=request.Form("ciudad")%>">
		<input type="hidden" name="pais" value="<%=request.Form("pais")%>">
		<input type="hidden" name="telefono" value="<%=request.Form("telefono")%>">
		<input type="hidden" name="email" value="<%=request.Form("email")%>">
		
		<input type="hidden" name="nombre_envio" value="<%=request.Form("nombre_envio")%>">
		<input type="hidden" name="apellidos_envio" value="<%=request.Form("apellidos_envio")%>">
		<input type="hidden" name="calle_envio" value="<%=request.Form("calle_envio")%>">
		<input type="hidden" name="numero_envio" value="<%=request.Form("numero_envio")%>">
		<input type="hidden" name="cp_envio" value="<%=request.Form("cp_envio")%>">
		<input type="hidden" name="ciudad_envio" value="<%=request.Form("ciudad")%>">
		<input type="hidden" name="pais_envio" value="<%=request.Form("pais_envio")%>">
		<input type="hidden" name="telefono_envio" value="<%=request.Form("telefono_envio")%>">
		<input type="hidden" name="email_envio" value="<%=request.Form("email_envio")%>">

	<div align="right"><p>
	<%if gastosenvio>0 then%>
	<input type="submit" value="Pagar">
	<%else%>
	<input type="submit" value="Realizar pedido">
	<%end if%>
	</p></div>
</form>
	<%

case "confirm"

	nif = ""& request.Form("nif")
	emailcopia = ""&request.Form("emailcopia")
	referencia = ""& request.Form("referencia")
	emailemision = ""& request.Form("emailemision")
	nombreemailemision = ""& request.Form("nombreemailemision")
	emailrecepcion = ""& request.Form("emailrecepcion")
	nombreemailrecepcion = ""& request.Form("nombreemailrecepcion")
	formapago = request.Form("formapago")
'	iva = request.Form("iva")
	gastosenvio = request.Form("gastosenvio")
	total = request.Form("total")

	' Enviar email a la central.
	'---------------------------------------------------------------------------------------------
	Subject = "Pedido: " & referencia
	Body = ""&session("str_pedido")
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
	
		titulo = referencia &" - "& nif
		seccion = 1
		seccion2 = 1
		usuario = numero(session("usuario"))	' ID del usuario logeado (ej: comercial que ordena el pedido)
		fuente = ""
		alfinal = 0
		enportada = 0
		activo = 1
		enlace = ""
		fecha = Date()					' Fecha en la que realiza el pedido
		fechaini = ""
		fechafin = ""
		resto_nombres = ", R_MEMO1, R_REF"
		resto_valores = ",'"& Body &"', '"& referencia &"'"
		conn = ruta_conn_pedidos
	
		str_insertar = ""& InsertarRegistro (titulo, seccion, seccion2, usuario, fuente, alfinal, enportada, activo, enlace, fecha, fechaini, fechafin, resto_nombres, resto_valores, conn)
	end if
	'---------------------------------------------------------------------------------------------

	if gastosenvio >0 then
		if formapago = "transferencia" then
			Response.Redirect("index.asp?secc=/transferencia")
		elseif formapago = "tarjeta" then
			Response.Redirect("http://www.mantones.com/banesto/cgi/totalizacion.exe?coste="& replace(euros(total),",",".") &"&moneda=EUR&nombre_comercio=CANDIDO_PUERTO&referencia="& referencia)
		else
			Response.Write "<b>Error:</b><br> No se ha especificado una forma de pago."	
		end if
	else
		Response.Redirect("index.asp?secc=/mensajes/pedidoconsultar")
	end if
	
case "datped"%>

	<b>Datos del ciente</b><br />
	<form action="index.asp?secc=/carrito&ac=vistafinal" method="post" name="f" onSubmit="return validar()">
		<input type="hidden" name="idi" value="<%=idioma%>"/>
		<input type="hidden" name="cualid" value="<%=cualid%>"/>
		<input type="hidden" name="seccrefer" value="<%=request.QueryString("seccrefer")%>"/>
		<input type="hidden" name="seccion" value="<%=request.QueryString("seccion")%>"/>
		<input type="hidden" name="seccion2" value="<%=request.QueryString("seccion2")%>"/>
		<%if ""& request.QueryString("msg") <> "" then
			Response.Write "<p><div align=center><em>"& request.QueryString("msg") &"</em></div></p>"
		end if%>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
    <TBODY>
      <TR>
        <TD colspan="3" align=left vAlign=top with="300"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><fieldset>
              <legend>Forma de pago </legend>
<input name="formapago" id="pagotarjeta" type="radio" value="tarjeta" checked>
              <label for="pagotarjeta">Tarjeta</label> <input name="formapago" id="pagotransferencia" type="radio" value="transferencia">
              <label for="pagotransferencia">Transferencia</label>
            </fieldset></td>
          </tr>
        </table>
        <br></TD>
      </TR>
      <TR>
        <TD width="50%" align=left vAlign=top with="300"><fieldset>
        <legend>Datos de facturaci&oacute;n</legend>
        <table width="100%"  border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td><table width="100%"  border="0" cellpadding="1" cellspacing="0">
                <tr align="left">
                  <td>Nombre</td>
                </tr>
                <tr align="left">
                  <td><input name=nombre type="text" class="campo" id="nombre" style="width:100%;" size=35 maxlength=45></td>
                </tr>
                <tr align="left">
                  <td>Apellidos</td>
                </tr>
                <tr align="left">
                  <td><input name=apellidos type="text" class="campo" id="apellidos" style="width:100%;" 
size=35 maxlength=45></td>
                </tr>
                <tr align="left">
                  <td>NIF</td>
                </tr>
                <tr align="left">
                  <td><input name=nif type="text" class="campo" id="nif" style="width:100%;" 
size=35 maxlength=45></td>
                </tr>
              </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Calle</td>
                    <td>&nbsp;</td>
                    <td>N&uacute;mero</td>
                  </tr>
                  <tr align="left">
                    <td><input name=calle type="text" class="campo" id="calle" 
size=20 maxlength=40></td>
                    <td>&nbsp;</td>
                    <td><input name=numero id="numero" 
size=10 maxlength=15 class="campo"></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>C&oacute;digo postal </td>
                    <td>&nbsp;</td>
                    <td>Ciudad</td>
                  </tr>
                  <tr align="left">
                    <td><input name=cp id="cp" 
size=20 maxlength=20 class="campo"></td>
                    <td>&nbsp;</td>
                    <td><input name=ciudad id="ciudad" 
size=20 maxlength=35 class="campo"></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Pais</td>
                  </tr>
                  <tr align="left">
                    <td><select name="pais" class="campo" id="pais" style="width:100%;">
                        <option value="AF">Afganist&aacute;n 
                        <option value="AL">Albania 
                        <option value="DE">Alemania 
                        <option value="AD">Andorra 
                        <option value="AO">Angola 
                        <option value="AI">Anguilla 
                        <option value="AQ">Ant&aacute;rtida 
                        <option value="AG">Antigua y Barbuda 
                        <option value="AN">Antillas Holandesas 
                        <option value="SA">Arabia Saud&iacute;
                        <option value="DZ">Argelia 
                        <option value="AR">Argentina 
                        <option value="AM">Armenia 
                        <option value="AW">Aruba 
                        <option value="AU">Australia 
                        <option value="AT">Austria 
                        <option value="AZ">Azerbaiy&aacute;n 
                        <option value="BS">Bahamas 
                        <option value="BH">Bahrein 
                        <option value="BD">Bangladesh 
                        <option value="BB">Barbados 
                        <option value="BE">B&eacute;lgica 
                        <option value="BZ">Belice 
                        <option value="BJ">Benin 
                        <option value="BM">Bermudas 
                        <option value="BY">Bielorrusia 
                        <option value="MM">Birmania 
                        <option value="BO">Bolivia 
                        <option value="BA">Bosnia y Herzegovina 
                        <option value="BW">Botswana 
                        <option value="BR">Brasil 
                        <option value="BN">Brunei 
                        <option value="BG">Bulgaria 
                        <option value="BF">Burkina Faso 
                        <option value="BI">Burundi 
                        <option value="BT">But&aacute;n 
                        <option value="CV">Cabo Verde 
                        <option value="KH">Camboya 
                        <option value="CM">Camer&uacute;n 
                        <option value="CA">Canad&aacute;
                        <option value="TD">Chad 
                        <option value="CL">Chile 
                        <option value="CN">China 
                        <option value="CY">Chipre 
                        <option value="VA">Ciudad del Vaticano (Santa Sede) 
                        <option value="CO">Colombia 
                        <option value="KM">Comores 
                        <option value="CG">Congo 
                        <option value="CD">Congo, Rep. Dem. del 
                        <option value="KR">Corea 
                        <option value="KP">Corea del Norte 
                        <option value="CI">Costa de Marf&iacute;l 
                        <option value="CR">Costa Rica 
                        <option value="HR">Croacia (Hrvatska) 
                        <option value="CU">Cuba 
                        <option value="DK">Dinamarca 
                        <option value="DJ">Djibouti 
                        <option value="DM">Dominica 
                        <option value="EC">Ecuador 
                        <option value="EG">Egipto 
                        <option value="SV">El Salvador 
                        <option value="AE">Emiratos &Aacute;rabes Unidos 
                        <option value="ER">Eritrea 
                        <option value="SI">Eslovenia 
                        <option value="ES" selected>Espa&ntilde;a 
                        <option value="US">Estados Unidos 
                        <option value="EE">Estonia 
                        <option value="ET">Etiop&iacute;a 
                        <option value="FJ">Fiji 
                        <option value="PH">Filipinas 
                        <option value="FI">Finlandia 
                        <option value="FR">Francia 
                        <option value="GA">Gab&oacute;n 
                        <option value="GM">Gambia 
                        <option value="GE">Georgia 
                        <option value="GS">Islas Sandwich
                        <option value="GH">Ghana 
                        <option value="GI">Gibraltar 
                        <option value="GD">Granada 
                        <option value="GR">Grecia 
                        <option value="GL">Groenlandia 
                        <option value="GP">Guadalupe 
                        <option value="GU">Guam 
                        <option value="GT">Guatemala 
                        <option value="GY">Guayana 
                        <option value="GF">Guayana Francesa 
                        <option value="GN">Guinea 
                        <option value="GQ">Guinea Ecuatorial 
                        <option value="GW">Guinea-Bissau 
                        <option value="HT">Hait&iacute;
                        <option value="HN">Honduras 
                        <option value="HK">Hong Kong, ZAE de la RPC 
                        <option value="HU">Hungr&iacute;a 
                        <option value="IN">India 
                        <option value="ID">Indonesia 
                        <option value="IQ">Irak 
                        <option value="IR">Ir&aacute;n 
                        <option value="IE">Irlanda 
                        <option value="BV">Isla Bouvet 
                        <option value="CX">Isla de Christmas 
                        <option value="IS">Islandia 
                        <option value="KY">Islas Caim&aacute;n 
                        <option value="CK">Islas Cook 
                        <option value="CC">Islas de Cocos o Keeling 
                        <option value="FO">Islas Faroe 
                        <option value="HM">Islas Heard y McDonald 
                        <option value="FK">Islas Malvinas 
                        <option value="MP">Islas Marianas del Norte 
                        <option value="MH">Islas Marshall 
                        <option value="UM">Islas menores de Estados Unidos 
                        <option value="PW">Islas Palau 
                        <option value="SB">Islas Salom&oacute;n 
                        <option value="SJ">Islas Svalbard y Jan Mayen 
                        <option value="TK">Islas Tokelau 
                        <option value="TC">Islas Turks y Caicos 
                        <option value="VI">Islas V&iacute;rgenes (EE.UU.) 
                        <option value="VG">Islas V&iacute;rgenes (Reino Unido) 
                        <option value="WF">Islas Wallis y Futuna 
                        <option value="IL">Israel 
                        <option value="IT">Italia 
                        <option value="JM">Jamaica 
                        <option value="JP">Jap&oacute;n 
                        <option value="JO">Jordania 
                        <option value="KZ">Kazajist&aacute;n 
                        <option value="KE">Kenia 
                        <option value="KG">Kirguizist&aacute;n 
                        <option value="KI">Kiribati 
                        <option value="KW">Kuwait 
                        <option value="LA">Laos 
                        <option value="LS">Lesotho 
                        <option value="LV">Letonia 
                        <option value="LB">L&iacute;bano 
                        <option value="LR">Liberia 
                        <option value="LY">Libia 
                        <option value="LI">Liechtenstein 
                        <option value="LT">Lituania 
                        <option value="LU">Luxemburgo 
                        <option value="MO">Macao 
                        <option value="MK">Macedonia
                        <option value="MG">Madagascar 
                        <option value="MY">Malasia 
                        <option value="MW">Malawi 
                        <option value="MV">Maldivas 
                        <option value="ML">Mal&iacute;
                        <option value="MT">Malta 
                        <option value="MA">Marruecos 
                        <option value="MQ">Martinica 
                        <option value="MU">Mauricio 
                        <option value="MR">Mauritania 
                        <option value="YT">Mayotte 
                        <option value="MX">M&eacute;xico 
                        <option value="FM">Micronesia 
                        <option value="MD">Moldavia 
                        <option value="MC">M&oacute;naco 
                        <option value="MN">Mongolia 
                        <option value="MS">Montserrat 
                        <option value="MZ">Mozambique 
                        <option value="NA">Namibia 
                        <option value="NR">Nauru 
                        <option value="NP">Nepal 
                        <option value="NI">Nicaragua 
                        <option value="NE">N&iacute;ger 
                        <option value="NG">Nigeria 
                        <option value="NU">Niue 
                        <option value="NF">Norfolk 
                        <option value="NO">Noruega 
                        <option value="NC">Nueva Caledonia 
                        <option value="NZ">Nueva Zelanda 
                        <option value="OM">Om&aacute;n 
                        <option value="NL">Pa&iacute;ses Bajos 
                        <option value="PA">Panam&aacute;
                        <option value="PG">Pap&uacute;a Nueva Guinea 
                        <option value="PK">Paquist&aacute;n 
                        <option value="PY">Paraguay 
                        <option value="PE">Per&uacute;
                        <option value="PN">Pitcairn 
                        <option value="PF">Polinesia Francesa 
                        <option value="PL">Polonia 
                        <option value="PT">Portugal 
                        <option value="PR">Puerto Rico 
                        <option value="QA">Qatar 
                        <option value="UK">Reino Unido 
                        <option value="CF">Rep&uacute;blica Centroafricana 
                        <option value="CZ">Rep&uacute;blica Checa 
                        <option value="ZA">Rep&uacute;blica de Sud&aacute;frica 
                        <option value="DO">Rep&uacute;blica Dominicana 
                        <option value="SK">Rep&uacute;blica Eslovaca 
                        <option value="RE">Reuni&oacute;n 
                        <option value="RW">Ruanda 
                        <option value="RO">Rumania 
                        <option value="RU">Rusia 
                        <option value="KN">Saint Kitts y Nevis 
                        <option value="WS">Samoa 
                        <option value="AS">Samoa Americana 
                        <option value="SM">San Marino 
                        <option value="VC">San Vicente y Granadinas 
                        <option value="SH">Santa Helena 
                        <option value="LC">Santa Luc&iacute;a 
                        <option value="ST">Santo Tom&eacute; y Pr&iacute;ncipe 
                        <option value="SN">Senegal 
                        <option value="SC">Seychelles 
                        <option value="SL">Sierra Leona 
                        <option value="SG">Singapur 
                        <option value="SY">Siria 
                        <option value="SO">Somalia 
                        <option value="LK">Sri Lanka 
                        <option value="PM">St. Pierre y Miquelon 
                        <option value="SZ">Suazilandia 
                        <option value="SD">Sud&aacute;n 
                        <option value="SE">Suecia 
                        <option value="CH">Suiza 
                        <option value="SR">Surinam 
                        <option value="TH">Tailandia 
                        <option value="TW">Taiw&aacute;n 
                        <option value="TZ">Tanzania 
                        <option value="TJ">Tayikist&aacute;n 
                        <option value="IO">Territorios brit&aacute;nicos del oc&eacute;ano &Iacute;ndico 
                        <option value="TF">Territorios franceses del Sur 
                        <option value="TP">Timor Oriental 
                        <option value="TG">Togo 
                        <option value="TO">Tonga 
                        <option value="TT">Trinidad y Tobago 
                        <option value="TN">T&uacute;nez 
                        <option value="TM">Turkmenist&aacute;n 
                        <option value="TR">Turqu&iacute;a 
                        <option value="TV">Tuvalu 
                        <option value="UA">Ucrania 
                        <option value="UG">Uganda 
                        <option value="UY">Uruguay 
                        <option value="UZ">Uzbekist&aacute;n 
                        <option value="VU">Vanuatu 
                        <option value="VE">Venezuela 
                        <option value="VN">Vietnam 
                        <option value="YE">Yemen 
                        <option value="YU">Yugoslavia 
                        <option value="ZM">Zambia 
                        <option value="ZW">Zimbabue 
                        </select>
                    </td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Tel&eacute;fono</td>
                  </tr>
                  <tr align="left">
                    <td><input name=telefono type="text" class="campo" id="telefono" style="width:100%;" 
size=20 maxlength=25></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Correo electr&oacute;nico</td>
                  </tr>
                  <tr align="left">
                    <td><input name=email type="text" class="campo" id="email" style="width:100%;" 
size=35 maxlength=40></td>
                  </tr>
              </table></td>
          </tr>
        </table>
        </fieldset>
        </TD>
        <TD align=left vAlign=top with="300">&nbsp;</TD>
        <TD width="50%" align=left vAlign=top with="300">          <fieldset>
        <legend>Datos de envio</legend>
        <table width="100%"  border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td><font color="#CC0000"><i>Rellenar  si es distinta a la de facturaci&oacute;n</i></font><br>
                <br>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Nombre</td>
                  </tr>
                  <tr align="left">
                    <td><input name=nombre_envio type="text" class="campo" id="nombre_envio" style="width:100%;" 
size=40 maxlength=45></td>
                  </tr>
                  <tr align="left">
                    <td>Apellidos</td>
                  </tr>
                  <tr align="left">
                    <td><input name=apellidos_envio type="text" class="campo" id="apellidos_envio" style="width:100%;" 
size=40 maxlength=45></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Calle</td>
                    <td>&nbsp;</td>
                    <td>N&uacute;mero</td>
                  </tr>
                  <tr align="left">
                    <td><input name=calle_envio type="text" class="campo" id="calle_envio" 
size=20 maxlength=40></td>
                    <td>&nbsp;</td>
                    <td><input name=numero_envio id="numero_envio" 
size=10 maxlength=15 class="campo"></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>C&oacute;digo postal </td>
                    <td>&nbsp;</td>
                    <td>Ciudad</td>
                  </tr>
                  <tr align="left">
                    <td><input name=cp_envio id="cp_envio" 
size=20 maxlength=20 class="campo"></td>
                    <td>&nbsp;</td>
                    <td><input name=ciudad_envio id="ciudad_envio" 
size=20 maxlength=35 class="campo"></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Pais</td>
                  </tr>
                  <tr align="left">
                    <td><select name="pais_envio" class="campo" id="pais_envio" style="width:100%;">
                      <option value="AF">Afganist&aacute;n 
                        <option value="AL">Albania 
                        <option value="DE">Alemania 
                        <option value="AD">Andorra 
                        <option value="AO">Angola 
                        <option value="AI">Anguilla 
                        <option value="AQ">Ant&aacute;rtida 
                        <option value="AG">Antigua y Barbuda 
                        <option value="AN">Antillas Holandesas 
                        <option value="SA">Arabia Saud&iacute;
                        <option value="DZ">Argelia 
                        <option value="AR">Argentina 
                        <option value="AM">Armenia 
                        <option value="AW">Aruba 
                        <option value="AU">Australia 
                        <option value="AT">Austria 
                        <option value="AZ">Azerbaiy&aacute;n 
                        <option value="BS">Bahamas 
                        <option value="BH">Bahrein 
                        <option value="BD">Bangladesh 
                        <option value="BB">Barbados 
                        <option value="BE">B&eacute;lgica 
                        <option value="BZ">Belice 
                        <option value="BJ">Benin 
                        <option value="BM">Bermudas 
                        <option value="BY">Bielorrusia 
                        <option value="MM">Birmania 
                        <option value="BO">Bolivia 
                        <option value="BA">Bosnia y Herzegovina 
                        <option value="BW">Botswana 
                        <option value="BR">Brasil 
                        <option value="BN">Brunei 
                        <option value="BG">Bulgaria 
                        <option value="BF">Burkina Faso 
                        <option value="BI">Burundi 
                        <option value="BT">But&aacute;n 
                        <option value="CV">Cabo Verde 
                        <option value="KH">Camboya 
                        <option value="CM">Camer&uacute;n 
                        <option value="CA">Canad&aacute;
                        <option value="TD">Chad 
                        <option value="CL">Chile 
                        <option value="CN">China 
                        <option value="CY">Chipre 
                        <option value="VA">Ciudad del Vaticano (Santa Sede) 
                        <option value="CO">Colombia 
                        <option value="KM">Comores 
                        <option value="CG">Congo 
                        <option value="CD">Congo, Rep. Dem. del 
                        <option value="KR">Corea 
                        <option value="KP">Corea del Norte 
                        <option value="CI">Costa de Marf&iacute;l 
                        <option value="CR">Costa Rica 
                        <option value="HR">Croacia (Hrvatska) 
                        <option value="CU">Cuba 
                        <option value="DK">Dinamarca 
                        <option value="DJ">Djibouti 
                        <option value="DM">Dominica 
                        <option value="EC">Ecuador 
                        <option value="EG">Egipto 
                        <option value="SV">El Salvador 
                        <option value="AE">Emiratos &Aacute;rabes Unidos 
                        <option value="ER">Eritrea 
                        <option value="SI">Eslovenia 
                        <option value="ES" selected>Espa&ntilde;a 
                        <option value="US">Estados Unidos 
                        <option value="EE">Estonia 
                        <option value="ET">Etiop&iacute;a 
                        <option value="FJ">Fiji 
                        <option value="PH">Filipinas 
                        <option value="FI">Finlandia 
                        <option value="FR">Francia 
                        <option value="GA">Gab&oacute;n 
                        <option value="GM">Gambia 
                        <option value="GE">Georgia 
                        <option value="GS">Islas Sandwich
                        <option value="GH">Ghana 
                        <option value="GI">Gibraltar 
                        <option value="GD">Granada 
                        <option value="GR">Grecia 
                        <option value="GL">Groenlandia 
                        <option value="GP">Guadalupe 
                        <option value="GU">Guam 
                        <option value="GT">Guatemala 
                        <option value="GY">Guayana 
                        <option value="GF">Guayana Francesa 
                        <option value="GN">Guinea 
                        <option value="GQ">Guinea Ecuatorial 
                        <option value="GW">Guinea-Bissau 
                        <option value="HT">Hait&iacute;
                        <option value="HN">Honduras 
                        <option value="HK">Hong Kong, ZAE de la RPC 
                        <option value="HU">Hungr&iacute;a 
                        <option value="IN">India 
                        <option value="ID">Indonesia 
                        <option value="IQ">Irak 
                        <option value="IR">Ir&aacute;n 
                        <option value="IE">Irlanda 
                        <option value="BV">Isla Bouvet 
                        <option value="CX">Isla de Christmas 
                        <option value="IS">Islandia 
                        <option value="KY">Islas Caim&aacute;n 
                        <option value="CK">Islas Cook 
                        <option value="CC">Islas de Cocos o Keeling 
                        <option value="FO">Islas Faroe 
                        <option value="HM">Islas Heard y McDonald 
                        <option value="FK">Islas Malvinas 
                        <option value="MP">Islas Marianas del Norte 
                        <option value="MH">Islas Marshall 
                        <option value="UM">Islas menores de Estados Unidos 
                        <option value="PW">Islas Palau 
                        <option value="SB">Islas Salom&oacute;n 
                        <option value="SJ">Islas Svalbard y Jan Mayen 
                        <option value="TK">Islas Tokelau 
                        <option value="TC">Islas Turks y Caicos 
                        <option value="VI">Islas V&iacute;rgenes (EE.UU.) 
                        <option value="VG">Islas V&iacute;rgenes (Reino Unido) 
                        <option value="WF">Islas Wallis y Futuna 
                        <option value="IL">Israel 
                        <option value="IT">Italia 
                        <option value="JM">Jamaica 
                        <option value="JP">Jap&oacute;n 
                        <option value="JO">Jordania 
                        <option value="KZ">Kazajist&aacute;n 
                        <option value="KE">Kenia 
                        <option value="KG">Kirguizist&aacute;n 
                        <option value="KI">Kiribati 
                        <option value="KW">Kuwait 
                        <option value="LA">Laos 
                        <option value="LS">Lesotho 
                        <option value="LV">Letonia 
                        <option value="LB">L&iacute;bano 
                        <option value="LR">Liberia 
                        <option value="LY">Libia 
                        <option value="LI">Liechtenstein 
                        <option value="LT">Lituania 
                        <option value="LU">Luxemburgo 
                        <option value="MO">Macao 
                        <option value="MK">Macedonia
                        <option value="MG">Madagascar 
                        <option value="MY">Malasia 
                        <option value="MW">Malawi 
                        <option value="MV">Maldivas 
                        <option value="ML">Mal&iacute;
                        <option value="MT">Malta 
                        <option value="MA">Marruecos 
                        <option value="MQ">Martinica 
                        <option value="MU">Mauricio 
                        <option value="MR">Mauritania 
                        <option value="YT">Mayotte 
                        <option value="MX">M&eacute;xico 
                        <option value="FM">Micronesia 
                        <option value="MD">Moldavia 
                        <option value="MC">M&oacute;naco 
                        <option value="MN">Mongolia 
                        <option value="MS">Montserrat 
                        <option value="MZ">Mozambique 
                        <option value="NA">Namibia 
                        <option value="NR">Nauru 
                        <option value="NP">Nepal 
                        <option value="NI">Nicaragua 
                        <option value="NE">N&iacute;ger 
                        <option value="NG">Nigeria 
                        <option value="NU">Niue 
                        <option value="NF">Norfolk 
                        <option value="NO">Noruega 
                        <option value="NC">Nueva Caledonia 
                        <option value="NZ">Nueva Zelanda 
                        <option value="OM">Om&aacute;n 
                        <option value="NL">Pa&iacute;ses Bajos 
                        <option value="PA">Panam&aacute;
                        <option value="PG">Pap&uacute;a Nueva Guinea 
                        <option value="PK">Paquist&aacute;n 
                        <option value="PY">Paraguay 
                        <option value="PE">Per&uacute;
                        <option value="PN">Pitcairn 
                        <option value="PF">Polinesia Francesa 
                        <option value="PL">Polonia 
                        <option value="PT">Portugal 
                        <option value="PR">Puerto Rico 
                        <option value="QA">Qatar 
                        <option value="UK">Reino Unido 
                        <option value="CF">Rep&uacute;blica Centroafricana 
                        <option value="CZ">Rep&uacute;blica Checa 
                        <option value="ZA">Rep&uacute;blica de Sud&aacute;frica 
                        <option value="DO">Rep&uacute;blica Dominicana 
                        <option value="SK">Rep&uacute;blica Eslovaca 
                        <option value="RE">Reuni&oacute;n 
                        <option value="RW">Ruanda 
                        <option value="RO">Rumania 
                        <option value="RU">Rusia 
                        <option value="KN">Saint Kitts y Nevis 
                        <option value="WS">Samoa 
                        <option value="AS">Samoa Americana 
                        <option value="SM">San Marino 
                        <option value="VC">San Vicente y Granadinas 
                        <option value="SH">Santa Helena 
                        <option value="LC">Santa Luc&iacute;a 
                        <option value="ST">Santo Tom&eacute; y Pr&iacute;ncipe 
                        <option value="SN">Senegal 
                        <option value="SC">Seychelles 
                        <option value="SL">Sierra Leona 
                        <option value="SG">Singapur 
                        <option value="SY">Siria 
                        <option value="SO">Somalia 
                        <option value="LK">Sri Lanka 
                        <option value="PM">St. Pierre y Miquelon 
                        <option value="SZ">Suazilandia 
                        <option value="SD">Sud&aacute;n 
                        <option value="SE">Suecia 
                        <option value="CH">Suiza 
                        <option value="SR">Surinam 
                        <option value="TH">Tailandia 
                        <option value="TW">Taiw&aacute;n 
                        <option value="TZ">Tanzania 
                        <option value="TJ">Tayikist&aacute;n 
                        <option value="IO">Territorios brit&aacute;nicos del oc&eacute;ano &Iacute;ndico 
                        <option value="TF">Territorios franceses del Sur 
                        <option value="TP">Timor Oriental 
                        <option value="TG">Togo 
                        <option value="TO">Tonga 
                        <option value="TT">Trinidad y Tobago 
                        <option value="TN">T&uacute;nez 
                        <option value="TM">Turkmenist&aacute;n 
                        <option value="TR">Turqu&iacute;a 
                        <option value="TV">Tuvalu 
                        <option value="UA">Ucrania 
                        <option value="UG">Uganda 
                        <option value="UY">Uruguay 
                        <option value="UZ">Uzbekist&aacute;n 
                        <option value="VU">Vanuatu 
                        <option value="VE">Venezuela 
                        <option value="VN">Vietnam 
                        <option value="YE">Yemen 
                        <option value="YU">Yugoslavia 
                        <option value="ZM">Zambia 
                        <option value="ZW">Zimbabue 
                    </select></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Tel&eacute;fono</td>
                  </tr>
                  <tr align="left">
                    <td><input name=telefono_envio type="text" class="campo" id="telefono_envio" style="width:100%;" 
size=20 maxlength=25></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                  <tr align="left">
                    <td>Correo electr&oacute;nico</td>
                  </tr>
                  <tr align="left">
                    <td><input name=email_envio type="text" class="campo" id="email_envio" style="width:100%;" 
size=35 maxlength=40></td>
                  </tr>
              </table></td>
          </tr>
        </table>
        </fieldset>
          <br>		    </TD>
      </TR>
      <TR>
        <TD align=left vAlign=top with="300">&nbsp;</TD>
        <TD align=left vAlign=top with="300">&nbsp;</TD>
        <TD align=right vAlign=top with="300"><input type="submit" name="Submit" value="Enviar"></TD>
      </TR>
    </TBODY>
  </TABLE>
	</form>
	<script language="javascript" type="text/javascript">
	<!--
		function validar(){
			if(f.nombre.value == ""){
				alert("Por favor, escriba su nombre.");
				f.nombre.focus()
				return false
			}
			if(f.apellidos.value == ""){
				alert("Por favor, escriba sus apellidos.");
				f.apellidos.focus()
				return false
			}
			if(f.nif.value == ""){
				alert("Por favor, escriba su nif.");
				f.nif.focus()
				return false
			}
			if(f.calle.value == ""){
				alert("Por favor, escriba su calle.");
				f.calle.focus()
				return false
			}
			if(f.numero.value == ""){
				alert("Por favor, escriba el numero de su calle.");
				f.numero.focus()
				return false
			}
			if(f.cp.value == ""){
				alert("Por favor, escriba su código postal.");
				f.cp.focus()
				return false
			}
			if(f.ciudad.value == ""){
				alert("Por favor, escriba su ciudad.");
				f.ciudad.focus()
				return false
			}
			if(f.telefono.value == ""){
				alert("Por favor, escriba su teléfono.");
				f.telefono.focus()
				return false
			}
			if(f.email.value == ""){
				alert("Por favor, escriba su correo electrónico.");
				f.email.focus()
				return false
			}
			return true
		}
	//-->
	</script>

<%case else

	' Actualizar datos
	' ------------------------------------------------------------------------------------------
	if ""& request.Form() <> "" then
		dim cItem, num_borrados, n, n2
		num_borrados = 0
		n=0
		for each cItem in carrito.nodo.childNodes

			idioma = cItem.getAttribute("idioma")
			cualid = cItem.getAttribute("cualidad")
			id = cItem.getAttribute("id")
			idx = cItem.getAttribute("idx") ' id única de xmlShop
			titulo = cItem.getAttribute("titulo")
			cantidadForm = numero(request.Form("cantidad_"& idx))
			cantidadActual = numero(cItem.getAttribute("cantidad"))

			' Borrar
			if bool(request.Form("borrar_"& idx)) then
				' Se borra
				'Response.Write "<br>Se borra"
				call carrito.delItem(idx)
			else
				if cantidadForm = 0 then
					' Se borra por cantidad
					'Response.Write "<br>Se borra por cantidad"
					call carrito.delItem(idx)
				elseif cantidadForm <> cantidadActual then
					' Hay un cambio
					'Response.Write "<br>Cambia cantidad de <b>"& titulo &"</b> de <b>"& cantidadActual &"</b> a <b>"& cantidadForm &"</b>"
					call carrito.addAtt(idx, "cantidad", cantidadForm)
				end if
			end if

		next
	else
		'
	end if

	if carrito.nodo.childNodes.length <= 0 then%>
		<b>Ning&uacute;n art&iacute;culo en carrito.</b>
	<%else%>
	<form name="f" action="index.asp?secc=/carrito" method="post">
		<input type="hidden" name="seccrefer" value="<%=request("seccrefer")%>">
		<input type="hidden" name="seccion" value="<%=request("seccion")%>">
		<input type="hidden" name="seccion2" value="<%=request("seccion2")%>">
		<table width="100%" border="0" cellpadding="2" cellspacing="0">
			<tr>
			<td><b>Artículo</b></td>
			<td align="center"><b>Cantidad</b></td>
			<td align="center"><b>Borrar</b></td>
			<td align="right"><b>Precio</b></td>
			</tr>
			<tr>
			<td colspan="4"><hr size="1" noshade></td>
			</tr>
			<%
			n=0
			for each item in carrito.nodo.childNodes
				n = n+1
				cantidad = numero(item.getAttribute("cantidad"))
				cantidad = numero(item.getAttribute("cantidad"))
				precio = numero(item.getAttribute("precio"))
				subtotalA = precio * cantidad
				subtotal = subtotal + subtotalA
				%>
				<tr>
				<td><span title="ID:<%=item.getAttribute("id")%>"><a href="index.asp?secc=<%=secc%>&ac=info&idi=<%=item.getAttribute("idioma")%>&cualid=<%=item.getAttribute("cualidad")%>&id=<%=item.getAttribute("id")%>&seccrefer=<%=request.QueryString("seccrefer")%>"><%=item.getAttribute("titulo")%></a></span></td>
				<td align="center"><input name="cantidad_<%=item.getAttribute("idx")%>" type="text" class="campo" value="<%=cantidad%>" size="2" maxlength="2"></td>
				<td align="center"><input type="checkbox" name="borrar_<%=item.getAttribute("idx")%>" value="1"></td>
				<td align="right"><%=euros(subtotalA)%> &euro;</td>
				</tr>
			<%next%>
		</table>
		<hr noshade size="1" />
		<br>
		<div align="right"><input type="submit" value="Actualizar datos">
		</div>
</form>
<div align="right">
		<b>Total:</b> <%=euros(subtotal)%> &euro;
		<br><br>
        <font color="#666666"><i>(No incluye IVA)</i></font></div>
		

	<%end if%>
	<br /><br />
	<div align="right">
	<a href="index.asp?secc=<%=request("seccrefer")%>&ac=listado&seccion=<%=request("seccion")%>&seccion2=<%=request("seccion2")%>"><img src="../esp/imagenes/seguir.gif" border="0"></a>
	<a href="index.asp?secc=/carrito&ac=datped&seccrefer=<%=request("seccrefer")%>&seccion=<%=request("seccion")%>&seccion2=<%=request("seccion2")%>"><img src="../esp/imagenes/formalizar.gif" border="0"></a></div>


	


    <%end select

if unerror then
	Response.Write "<b>Error:</b><br />"& msgerror
end if


		
%>

	