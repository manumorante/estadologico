<%Dim unerror, msgerror, rextotal, rex%>
<!--#include file="inc_conn.asp" -->
<!--#include virtual="/admin/global/inc_inicia_xml.asp" -->
<!--#include virtual="/admin/global/inc_rutinas.asp" -->
<!--#include virtual="/admin/inc_sha256.asp" -->

<%if unerror then%>
	<b>Error</b><br><%=msgerror%>
<%else

	inicia_xml
	
	if ""&idioma = "" then
		pi = request.ServerVariables("PATH_INFO")
		if inStr(pi,"/esp/index.asp") then
			idioma = "esp"
		elseif inStr(pi,"/eng/index.asp") then
			idioma = "eng"
		elseif inStr(pi,"/fra/index.asp") then
			idioma = "fra"
		elseif inStr(pi,"/deu/index.asp") then
			idioma = "deu"
		elseif inStr(pi,"/ita/index.asp") then
			idioma = "ita"
		else
			if session("idioma") <> "" then
				idioma = session("idioma")
			else
				unerror = true : msgerror = "No ha sido posible determinar el idioma de navegación / aSkipper"
			end if
		end if
	end if
	
	

	' Ac
	ac = ""&request("ac")
	id = ""&request.QueryString("id")

	' Paginado
	registrosPorPagina = 5

	'Inicio (Cabeza lectora)
	Dim pag
	pag = ""&request("pag")
	if pag <> "" then
		if not esNumero(pag) then
			pag = 1
		end if
	else
		pag = 1
	end if

	' Sección 
	Dim seccion
	if esNumero(request("seccion")) then
		seccion = numero(request("seccion"))
	else
		seccion = -1
	end if
	
	' Sección2
	Dim seccion2
	if esNumero(request("seccion2")) then
		seccion2 = request("seccion2")
	else
		seccion2 = -1
	end if
	
	' Cadena de búsqueda
	Dim cadena
	if ""&request("cadena") <> "" then
		cadena = request("cadena")
	else
		cadena = ""
	end if



%>

<script language="JavaScript" type="text/javascript">
	<!--
	
	function ir(q){
		var q_seccion = ""
		if (q.indexOf("&seccion=")<0 ){
			q_seccion = "&seccion=<%=seccion%>"
		}
		var q_pag = ""
		if (q.indexOf("&pag=")<0 ){
			q_pag = "&pag=<%=pag%>"
		}
		location.href="index.asp?secc=<%=request.QueryString("secc")%>"+q+q_seccion+q_pag
	}
	// envio
	function envio() {
		ir("&cadena="+f.cadena.value)
		return false
	}
	// todo
	function todo() {
		f.ac.value = ""
		f.cadena.value = ""
		f.seccion.value = ""
		f.cadenaanterior.value = ""
		f.pag.value = 1
		f.submit()
	}
	function buscar() {
		ir("&cadena="+f.cadena.value)
	}
	function irPag(num) {
		ir("&pag="+num)
	}
	function ampliar(id) {
		ir("&ac=ampliar&id="+id)
	}
	function irSeccion(seccion, seccion2) {
		ir("&seccion="+seccion+"&seccion2="+seccion2+"&pag=1")
	}
	function todasSecciones() {
		ir("&seccion=&seccion2=")
	}
	function todasSecciones2() {
		ir("&seccion2=")
	}
	//
	// volver
	function volver() {
		location="index.asp?secc=<%=secc%>"
	}
	//-->
</script>
<%

	

	' Cargar las secciones a un vector --
	sql = ""
	sql = sql & "SELECT *"
	sql = sql & " FROM SECCIONES"
	sql = sql & " WHERE S_REGISTROS > 0"
	sql = sql & " ORDER BY S_ORDEN "
	set reseccion = Server.CreateObject("ADODB.Recordset")
	on error resume next
		reseccion.ActiveConnection = conn_
		if err<>0 then
			unerror = true : msgerror = "No se ha podido cargar la base de datos para "&ucase(cualid)&" en el idioma "& ucase(idioma) &".<br>Conpruebe que la base de datos esté disponible en el idioma seleccionado."
		else
			reseccion.Source = sql : reseccion.CursorType = 3 : reseccion.CursorLocation = 2 : reseccion.LockType = 3
			reseccion.Open()
			if err<>0 then
				unerror = true : msgerror = "La consulta a la base de datos contiene algún error.<br>SQL:["&sql&"]"
			else
				Dim arrSeccionesId()
				Dim arrSeccionesNombre()
				Dim arrSeccionesTipo()
				if not 	reseccion.eof then
					numSecciones = 0
					while not reseccion.eof
						redim preserve arrSeccionesId(numSecciones) : arrSeccionesId(numSecciones) = reseccion("S_ID")
						redim preserve arrSeccionesNombre(numSecciones) : arrSeccionesNombre(numSecciones) = reseccion("S_NOMBRE")
						numSecciones = numSecciones + 1
						reseccion.movenext
					wend
				end if
			end if
		end if
		reseccion.Close()
		set reseccion = Nothing
	on error goto 0


	
	if not unerror then
	
		' Resetear paginado (pag=1) al hacer nueva búsqueda
		if StrComp(request.Form("cadenaanterior"),cadena) <>0 then
			pag = 1
		end if
		
		' CONSULTA ----------------------------------------------------------------------------------------------------
		sql = ""
		sql = sql & "SELECT *"
		sql = sql & " FROM REGISTROS, SECCIONES"
		sql = sql & " WHERE R_SECCION = S_ID"
		
		' Foro
		sql = sql & " AND R_ID_HIJODE = 0"		

		' Búsqueda
		if cadena <> "" then
			sql = sql & " AND (R_TITULO LIKE '%"& Replace(cadena,"'","''") & "%' OR R_MEMO1 LIKE '%"& Replace(cadena,"'","''") &"%')"
		end if
		
		if nav_activo then
			sql = sql & " AND R_ACTIVO = 1"		
		end if
	
		' Sección filtra
	
		if nav_filtraralias <> "" then
			if inStr(nav_filtraralias,",") then
			else
				sql = sql & " AND S_ALIAS = '"& nav_filtraralias &"'"
			end if
		end if
	
		if nav_activo_secciones then
			if numero(seccion) <> -1 then
				sql = sql & " AND R_SECCION = "& seccion
			end if
		end if
		if nav_activo_secciones then
			if numero(seccion2) <> -1 then
				sql = sql & " AND R_SECCION2 = "& seccion2
			end if
		end if
		
		if nav_orden <> "" then
			sql = sql & " ORDER BY "& nav_orden &""
		else
			sql = sql & " ORDER BY R_ORDEN ASC"
		end if
	
		'Response.Write("<br><fon size=1>"&sql&"</font>")
		' -------------------------------------------------------------------------------------------------------------
	
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3
		re.Open()
		if not re.eof then
			re.move ((pag * registrosPorPagina)) - registrosPorPagina
			totalRegistros = re.recordcount		
		else
			totalRegistros = 0
		end if

	end if ' unerror


if not unerror then%>
<!-- INICIO DE FORMULARIO --------------------------------------------------------------------------------------- -->
<input type="hidden" name="ac" value="<%=ac%>">
<input type="hidden" name="ref" value="">
<input type="hidden" name="id" value="">
<input type="hidden" name="pag" value="<%=pag%>">
<input type="hidden" name="seccion" value="<%=seccion%>">
<input type="hidden" name="seccion2" value="<%=seccion2%>">
<input type="hidden" name="seccion3" value="<%=seccion3%>">

<%if nav_buscar then%>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="noticias-busqueda">
	  <tr>
		<td>
		<table border="0" cellpadding="2" cellspacing="0">
			<tr>
			<td valign="middle">Texto a buscar: <input name="cadena" type="text" class="noticias-input" value="<%=cadena%>" size="30" maxlength="150"><input name="cadenaanterior" type="hidden" value="<%=cadena%>"></td>
			<td valign="bottom"><a href="JavaScript:buscar();"><img src="/<%=c_s%><%=idioma%>/imagenes/buscar.gif" border="0"></a><a href="JavaScript:todo();"><img src="/<%=c_s%><%=idioma%>/imagenes/vertodo.gif" border="0"></a></td>
			</tr>
		</table></td>
	  </tr>
  </table>
  <br>
<%else%>
  <input type="hidden" name="cadena" value="">
  <input name="cadenaanterior" type="hidden" value="<%=cadena%>">
<%end if%>


<%if cadena <> "" then
	num = 0
	numPaginas = 0
	for n=1 to totalRegistros
		if n mod registrosPorPagina = 1 then
			num = num +1
			numPaginas = numPaginas + 1
		end if
	next

%>
  <br>
  <table width="100%" border="0" cellpadding="4" cellspacing="0" bgcolor="f8f8f8">
    <tr><td><font color="#333333">Se han encontrado <b><%=totalRegistros%></b> resultados
	        en la b&uacute;squeda para &quot;<font color="#5555aa"><b><%=cadena%></b></font>&quot;.</font></td>
      <td align="right"> <font color="#333333">P&aacute;gina <%=pag%> de <%=numPaginas%></font>.</td>
	</tr>
  </table>
  <br>
  <%end if%>


<%select case ac

case "nuevo_tema"

inserta = false
if request.Form() <> "" then

	' Declaro el codigo de usuaro en 0. Usuario no registrado.
	cod_usuario = 0 

	' Comprobamos si ha insertado una nombre de usuario y clave. Si es asi lo validamos.
	c_usuario = "" & request.Form("usuario")
	c_clave = sha256("" & request.Form("clave"))
	if (c_usuario <> "" and c_clave <> "") or ""&session("usuario") <> "" then
		if c_clave = getClave(c_usuario) then
			cod_usuario = getCodigo(c_usuario)
			session("usuario") = cod_usuario
			inserta = true
		else
			if ""&request.Form("registrar") = "1" and ""&session("usuario") = "" then
			
				' Registrar usuario
				ruta_xml_usuarios = "/" & c_s & "datos/usuarios.xml"
				Set xml_usuarios = CreateObject("MSXML.DOMDocument")
				if not xml_usuarios.Load(server.MapPath(ruta_xml_usuarios)) then
					unerror = true : msgerror = "No se ha encontrado el XML de usuarios."
				else
					set nodo_usuarios = xml_usuarios.selectSingleNode("datos/usuarios")
					if not typeOK(nodo_usuarios) then
						unerror = true : msgerror = "El XML de usuario está corrupto."
					end if
				end if
				
				if not unerror then
					cod_usuario = numero(nodo_usuarios.getAttribute("maxcodigo")) + 1
					creaAtributo nodo_usuarios,"maxcodigo",cod_usuario
					set nuevo = xml_usuarios.createElement("usuario")
					creaAtributo nuevo,"fecharegistro",date()
					creaAtributo nuevo,"idioma",idioma
					creaAtributo nuevo,"grupo",cod_grupo
					creaAtributo nuevo,"codigo",cod_usuario
					creaAtributo nuevo,"usuario",c_usuario
					creaAtributo nuevo,"email",request.Form("email")
					creaAtributo nuevo,"clave",c_clave
					
					nodo_usuarios.appendChild(nuevo)
					xml_usuarios.save server.MapPath(ruta_xml_usuarios)
				end if
				inserta = true
			
			elseif ""&session("usuario") <> "" then
				cod_usuario = session("usuario")
				inserta = true
			else
				' Rebote de form
				%>
				<script>
				function confirmar(){
					if (f.clave.value != "" && f.clave2.value != f.clave.value) {
						alert("Las claves no coinciden.")
						f.clave2.focus()
						return false
					}
				}
				</script>
				<form name="f" method="post" action="index.asp?secc=<%=secc%>&ac=nuevo_tema" onSubmit="return confirmar()">
				<input type="hidden" name="registrar" value="1">
				<%for each a in request.Form()%>
					<input type=hidden name="<%=a%>" value="<%=request.Form(a)%>">
				<%next%>
				El usuario indicado no está registrado, confirme su clave e indique su email para completar el registro del mismo.
				<br>
				<br>
				<table  border="0" align="center" cellpadding="1" cellspacing="0">
				  <tr>
					<td align="right"><b>Usuario:</b></td>
					<td><%=c_usuario%></td>
				  </tr>
				  <tr>
					<td align="right"><b>Clave:</b></td>
					<td><input name="clave2" class="campo" type="password" id="clave2"></td>
				  </tr>
				  <tr>
					<td align="right"><b>E-mail:</b></td>
					<td><input name="email" class="campo" type="text" id="email"></td>
				  </tr>
				</table>
				<br>
				<br>
				<div align="center">
				  <input name="" type="button" class="campo" onClick="window.history.back()" value="Atrás">
				  <input type="submit" class="campo" name="Submit" value="Enviar">
				</div>
				<%
			end if
		end if
	else
		inserta = true
	end if%>
	
	</form>
	<%
	
	if inserta then
		' Insertar el regisro
		titulo = ""&request.Form("asunto")
		seccion = numero(request.Form("seccion"))
		seccion2 = numero(request.Form("seccion2"))
		usuario = cod_usuario
		fuente = ""
		alfinal = 0
		enportada = 0
		activo = 0
		enlace = ""
		fecha = ""
		fechaini = ""
		fechafin = ""
		resto_nombres = ", R_MEMO1"
		resto_valores = ", '"& request.Form("mensaje") &"'"
		conn = conn_
		call insertarRegistro(titulo, seccion, seccion2, usuario, fuente, alfinal, enportada, activo, enlace, fecha, fechaini, fechafin, resto_nombres, resto_valores, conn)
		
		' Incrementar contador
		call exeSQL("UPDATE REGISTROS SET R_NUM_RESPUESTAS = R_NUM_RESPUESTAS + 1 WHERE R_ID = "& request.Form("id"),conn_)

		Response.Redirect("index.asp?secc=" & secc	)		
	end if
	
	


	


else%>
	<br><b>Nuevo tema</b>
	
	<form name="f" action="index.asp?secc=<%=secc%>&ac=nuevo_tema" method="post" onSubmit="return enviar()">
	
	


	<script language="javascript" type="text/javascript">
		function enviar(){
			if(f.seccion.value == ""){
				alert("Por favor, elija una sección.")
				f.seccion.focus()
				return false
			}
			<%if ""&session("usuario") = "" then%>
			if(f.usuario.value != "" && f.clave.value == ""){
				if (confirm("Ha escrito un nombre de usuario.\n¿Desea escribir su clave de usuario registrado o proceder a su registro?.\nSi pulsa 'Cancelar' aparecerá como mensaje anónimo.")){
					f.clave.focus()
					return false
				}
			}
			<%end if%>
			if(f.asunto.value == ""){
				alert("Por favor, escriba el asunto de su respuesta.")
				f.asunto.focus()
				return false
			}
			if(f.mensaje.value == ""){
				alert("Por favor, escriba su mensaje.")
				f.mensaje.focus()
				return false
			}
		}
	</script>
	<table width="100%"  border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
	<tr>
	  <td>

	<%
	if nav_activo_secciones then

		' Secciones
		sql = "SELECT * FROM SECCIONES ORDER BY S_ORDEN"
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_ : re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3
		re.Open()
		if not re.eof then%>	  
			<table border="0" cellspacing="0" cellpadding="1">
				<tr>
				<td>
				<b>Secci&oacute;n:</b>
				<select name="seccion" class="campo">
				<option value="">[ELIJA UNA SECCIÓN]</option>
				<%while not re.eof%>
					<option value="<%=re("S_ID")%>" <%if seccion = re("S_ID") then Response.Write "selected" end if%>><%=re("S_NOMBRE")%></option>
				<%re.movenext : wend%>
				</select>
				</td>
				</tr>
			</table>
		<%else%>
			<input type="hidden" name="seccion" value="<%=seccion%>">
		<%end if ' not eof
		re.Close()
		set re = nothing
	else%>
		<input type="hidden" name="seccion" value="<%=seccion%>">
	<%end if ' activo secciones%>

	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
		<tr>
		<td height="5" align="right"><img src="../../spacer.gif" width="1" height="1"></td>
		</tr>
	</table>
	

<%if ""&session("usuario") = "" then%>
	<table  border="0" cellpadding="1" cellspacing="0">
              <tr>
                <td align="right"><b>Usuario:</b>
                    <input name="usuario" type="text" class="campo" id="usuario" size="15">
                    <b>Clave:</b>
                    <input name="clave" type="password" class="campo" id="clave" size="15"></td>
              </tr>
        </table>
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="5" align="right"><img src="../../spacer.gif" width="1" height="1"></td>
                </tr>
              </table>
<%end if%>  

              <table width="100%"  border="0" cellspacing="0" cellpadding="1">
                <tr>
                  <td align="right"><b>Asunto:</b></td>
                  <td width="100%"><input name="asunto" type="text" class="campo" id="asunto" style="width:100%"></td>
                </tr>
              </table>
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="5" align="right"><img src="../../spacer.gif" width="1" height="1"></td>
                </tr>
              </table>
              <table width="100%"  border="0" cellpadding="1" cellspacing="0">
                <tr>
                  <td width="100%"><b>Mensaje</b></td>
                </tr>
                <tr>
                  <td><textarea name="mensaje" cols="50" rows="10" wrap="virtual" class="campo" id="mensaje" style="width:100%"></textarea></td>
                </tr>
        </table></td></tr>
  </table>
  <br>
  <table width="100%"  border="0" align="center" cellpadding="1" cellspacing="0">
    <tr>
      <td align="right"><input name=""class="campo" type="button" onClick="window.history.back()" value="Volver">
        <input type="submit" class="campo" value="Enviar">
      </td>
    </tr>
  </table>
</form>
<br>
<br>
<%end if%>


<%

' Procedo unico para respuestas a temas principales o a otras respuestas 
case "responder"

if request.Form() <> "" then

	cod_usuario = 0
	if ""&session("usuario") <> "" then
		cod_usuario = session("usuario")
		inserta = true
	else
		c_usuario = "" & request.Form("usuario")
		c_clave = sha256("" & request.Form("clave"))
		if c_usuario = "" or c_clave = "" then
			inserta = true
		else
			if c_clave = getClave(c_usuario) then
				cod_usuario = getCodigo(c_usuario)
				session("usuario") = cod_usuario
				inserta = true
			else
			%>

<script>
				function confirmar(){
					if (f.clave.value != "" && f.clave2.value != f.clave.value) {
						alert("Las claves no coinciden.")
						f.clave2.focus()
						return false
					}
				}
				</script>
				<form name="f" method="post" action="index.asp?secc=<%=secc%>&ac=nuevo_tema" onSubmit="return confirmar()">
				<input type="hidden" name="registrar" value="1">
				<%for each a in request.Form()%>
					<input type=hidden name="<%=a%>" value="<%=request.Form(a)%>">
				<%next%>
				El usuario indicado no está registrado, confirme su clave e indique su email para completar el registro del mismo.
				<br>
				<br>
				<table  border="0" align="center" cellpadding="1" cellspacing="0">
				  <tr>
					<td align="right"><b>Usuario:</b></td>
					<td><%=c_usuario%></td>
				  </tr>
				  <tr>
					<td align="right"><b>Clave:</b></td>
					<td><input name="clave2" class="campo" type="password" id="clave2"></td>
				  </tr>
				  <tr>
					<td align="right"><b>E-mail:</b></td>
					<td><input name="email" class="campo" type="text" id="email"></td>
				  </tr>
				</table>
				<br>
				<br>
				<div align="center">
				  <input name="" type="button" class="campo" onClick="window.history.back()" value="Atrás">
				  <input type="submit" class="campo" name="Submit" value="Enviar">
				</div>

			<%
			end if
		end if ' if c_usuario = "" or c_clave = "" then
	end if
	
	if inserta then

		' Insertar el regisro
		titulo = ""&request.Form("asunto")
		seccion = numero(request.Form("seccion"))
		seccion2 = numero(request.Form("seccion2"))
		usuario = cod_usuario
		fuente = ""
		alfinal = 0
		enportada = 0
		activo = 0
		enlace = ""
		fecha = ""
		fechaini = ""
		fechafin = ""
		resto_nombres = ", R_MEMO1, R_ID_HIJODE"
		resto_valores = ", '"& request.Form("mensaje") &"', "& numero(request.Form("id"))
		conn = conn_
	
		call insertarRegistro(titulo, seccion, seccion2, usuario, fuente, alfinal, enportada, activo, enlace, fecha, fechaini, fechafin, resto_nombres, resto_valores, conn)
	
		' Incrementar contador
		call exeSQL("UPDATE REGISTROS SET R_NUM_RESPUESTAS = R_NUM_RESPUESTAS + 1 WHERE R_ID = "& request.Form("id"),conn_)
		
		Response.Redirect("index.asp?secc=/ocio/foro&ac=ampliar&id="& numero(request.Form("id")) &"&seccion=" & numero(request.Form("seccion")) & "&pag=1")
	end if

else

%>
<form name="f" action="index.asp?secc=<%=secc%>&ac=responder" method="post" onSubmit="return enviar()">
<input name="seccion" type="hidden" value="<%=request("seccion")%>">
<input name="id" type="hidden" value="<%=request("id")%>">
<script language="javascript" type="text/javascript">
	function enviar(){
			<%if ""&session("usuario") = "" then%>
			if(f.usuario.value != "" && f.clave.value == ""){
				if (confirm("Ha escrito un nombre de usuario.\n¿Desea escribir su clave de usuario registrado o proceder a su registro?.\nSi pulsa 'Cancelar' aparecerá como mensaje anónimo.")){
					f.clave.focus()
					return false
				}
			}
			<%end if%>

		if(f.asunto.value == ""){
			alert("Por favor, escriba el asunto de su respuesta.")
			f.asunto.focus()
			return false
		}
		if(f.mensaje.value == ""){
			alert("Por favor, escriba su mensaje.")
			f.mensaje.focus()
			return false
		}
	}
</script>
<table width="100%"  border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td>
	
	<%if ""&session("usuario") = "" then%>
	<table  border="0" cellpadding="1" cellspacing="0">
        <tr>
          <td align="right"><b>Usuario:</b>
              <input name="usuario" type="text" class="campo" id="usuario" size="15">
              <b>Clave:</b>
              <input name="clave" type="password" class="campo" id="clave" size="15"></td>
        </tr>
      </table>
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5" align="right"><img src="../../spacer.gif" width="1" height="1"></td>
          </tr>
        </table>
		
		<%end if%>
		
        <table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td align="right"><b>Asunto:</b></td>
            <td width="100%"><input name="asunto" type="text" class="campo" id="asunto" style="width:100%"></td>
          </tr>
        </table>
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5" align="right"><img src="../../spacer.gif" width="1" height="1"></td>
          </tr>
        </table>
        <table width="100%"  border="0" cellpadding="1" cellspacing="0">
          <tr>
            <td width="100%"><b>Mensaje</b></td>
          </tr>
          <tr>
            <td><textarea name="mensaje" cols="50" rows="10" wrap="virtual" class="campo" id="mensaje" style="width:100%"></textarea></td>
          </tr>
      </table></td>
  </tr>
</table>
<br>
<table width="100%"  border="0" align="center" cellpadding="1" cellspacing="0">
  <tr>
    <td align="right"><input name="" type="button" onClick="window.history.back()" value="Volver">
        <input type="submit" value="Enviar">
    </td>
  </tr>
</table>
</form>
<%
end if ' form
	

case "ampliar" ' --------------------------------------------------------------------------- Ampliar noticia

	id = request.QueryString("id")&request.Form("id")
	if ""&id = "" or not isNumeric(id) then
		Response.Redirect(request.ServerVariables("HTTP_REFERER")&"&No se ha encontrado el registro solicitado.")
	end if
	


	if id <> "" then
		sql = ""
		sql = sql & "SELECT *"
		sql = sql & " FROM REGISTROS"
		sql = sql & " WHERE R_ID = " & id
		set re = Server.CreateObject("ADODB.Recordset")
		re.ActiveConnection = conn_
		re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3
		re.Open()
		if re.eof then
			Response.Redirect(request.ServerVariables("HTTP_REFERER")&"&msg=No hay ningún registro con esta referencia.")
		else%>
	  <table border="0" cellpadding="6" cellspacing="0" width="100%" align="center">
		  <tr>
			<td>

			<%if nav_fecha and re("R_FECHA") <> 0 then%>
				<table border="0" align="right" cellpadding="0" cellspacing="0">
					<tr>
					<td><span class="noticias-fecha"><%=re("R_FECHA")%></span></td>
					</tr>
				</table>
		    <%end if%>
			
			<span class="noticias-titulo"><%=re("R_TITULO")%></span></td>

			

		
	  </table>
	  <%if nav_cuerpo <> "" then
				if re("R_"&nav_cuerpo) <> "" then%>
	  <table width="100%"  border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td><span class="noticias-subtitulo"><%=re("R_"&nav_cuerpo)%></span></td>
        </tr>
      </table>
		<%end if
			end if%>
              <table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#666666">
                <tr>
                  <td><table width="100%"  border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
                    <tr>
                      <td><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                        <tr>
                          <td><%if nav_foto and re("R_FOTO") <> "" then%>
                              <table border='0' align="right" cellpadding='0' cellspacing='0'>
                                <tr>
                                  <td><img src='/<%=c_s%>img/foto_s_i.gif'></td>
                                  <td background='/<%=c_s%>img/foto_s.gif'><img src='/<%=c_s%>spacer.gif' width='1' height='1'></td>
                                  <td><img src='/<%=c_s%>img/foto_s_d.gif'></td>
                                </tr>
                                <tr>
                                  <td background='/<%=c_s%>img/foto_i.gif'><img src='/<%=c_s%>spacer.gif' width='1' height='1'></td>
                                  <td><table cellpadding='0' cellspacing='0' border='0'>
                                      <tr>
                                        <td class='general-foto-pixel' align='center'><img src="/<%=c_s%>datos/<%=idioma%>/<%=cualid%>/fotos/<%=re("R_FOTO")%>"></td>
                                      </tr>
                                  </table></td>
                                  <td background='/<%=c_s%>img/foto_d.gif'><img src='/<%=c_s%>spacer.gif' width='1' height='1' border='0'></td>
                                </tr>
                                <tr>
                                  <td><img src='/<%=c_s%>img/foto_b_i.gif'></td>
                                  <td background='/<%=c_s%>img/foto_b.gif'><img src='/<%=c_s%>spacer.gif' width='1' height='1'></td>
                                  <td><img src='/<%=c_s%>img/foto_b_d.gif'></td>
                                </tr>
                              </table>
                              <%end if%>
                              <%=escribeHtml(re("R_MEMO1"))%>
                              
                          </td>
                        </tr>
                      </table></td>
                    </tr>
                  </table></td>
                </tr>
              </table>


      <%
		sql = "SELECT * FROM REGISTROS WHERE R_ID_HIJODE = " & id
		set respuestas = Server.CreateObject("ADODB.Recordset")
		respuestas.ActiveConnection = conn_
		respuestas.Source = sql : respuestas.CursorType = 1 : respuestas.CursorLocation = 2 : respuestas.LockType = 1
		respuestas.Open()%>
		
		<%if respuestas.eof then%>
		  <br>
		  <div align="center"><b>No hay respuestas para este tema</b></div>
		  <br>
		<%else%>
		  <br>
		  <b>RESPUESTAS</b>
		  <br>
		  <br>

		
			<%while not respuestas.eof%>
	      <table width="95%"  border="0" align="center" cellpadding="1" cellspacing="0">
  <tr>
    <td><img src="/<%=c_s%>img/ico_disco_duro.gif" align="absbottom">&nbsp;
	
	            <%if re("R_USUARIO") > 0 then%>
			  	<span title=" Código: <%=re("R_USUARIO")%> "><%=getNombreUsuario(re("R_USUARIO"))%></span>
			<%else%>
				An&oacute;nimo<%end if%>:
				<b><%=respuestas("R_TITULO")%></b></td>
    <td align="right"><span title=" <%=respuestas("R_AUTOHORA")%> "><font color="#666666" size="1"><%=respuestas("R_AUTOFECHA")%></font></span></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#ECE9D8"><img src="../../spacer.gif" width="1" height="1"></td>
    </tr>
  <tr>
    <td colspan="2"><table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#ECE9D8">
      <tr>
        <td><table width="100%"  border="0" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td><%=escribeHtml(respuestas("R_MEMO1"))%></td>
          </tr>
        </table></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
	      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="10"><img src="../../spacer.gif" width="1" height="1"></td>
            </tr>
</table>
	      <%
				respuestas.movenext()
			wend
			
			end if ' eof
			respuestas.Close()
		set respuestas = nothing
		%>
		
<br>

<table width="100%" border="0" cellpadding="6" cellspacing="0">
	    <tr>
	      <td align="right">		      <a href="javascript:volver()">Volver</a> -  <a href="index.asp?secc=<%=secc%>&ac=responder&id=<%=re("R_ID")%>&seccion=<%=re("R_SECCION")%>">Responder</a>      </td>
        </tr>
</table>
<%
		end if
		re.close
		set re = nothing
	end if

case else ' --------------------------------------------------------------------------- Listado, búsquedas ...

' Secciones
if nav_activo_secciones and int(numSecciones) > 1 then%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="noticias-secciones">

		<tr>
	<td width="60"  class="noticias-rotulosecciones">&nbsp;Secciones: </td>
		<td align="left"><table border="0" cellpadding="2" cellspacing="3">
		<tr>
		<%for n=0 to numSecciones-1

				if int(arrSeccionesId(n)) = cint(seccion) then%>
					<td class="boton-over">&nbsp;<%=arrSeccionesNombre(n)%>&nbsp;</td>
				<%else%>
					<td class="boton-out"><a href="JavaScript:irSeccion(<%=arrSeccionesId(n)%>,-1)">&nbsp;<%=arrSeccionesNombre(n)%>&nbsp;</a></td>
				<%end if

		next
			if int(seccion) = -1 then%>
				<td class="boton-over">&nbsp;Todas&nbsp;</td>
			<%else%>
				<td class="boton-out">&nbsp;<a href="JavaScript:todasSecciones()">Todas</a>&nbsp;</td>
			<%end if%>
		</tr>
		</table></td>
		</tr>
	</table>
<%end if%>
	<%if nav_activo_secciones2 and seccion <> -1 then%>
    <table  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="30">&nbsp;</td>
        <td><table border="0" cellpadding="0" cellspacing="0" class="noticias-secciones">
          <tr>
            <td width="60"  class="noticias-rotulosecciones">&nbsp;&nbsp;Subsecciones: </td>
            <td align="left"><%
	' Subsecciones
	sql = "SELECT * FROM SECCIONES2 WHERE S2_REGISTROS > 0 AND S2_ID_S = "& seccion &" ORDER BY S2_ORDEN "
	set re_subsecc = Server.CreateObject("ADODB.Recordset")
	re_subsecc.ActiveConnection = conn_
	re_subsecc.Source = sql : re_subsecc.CursorType = 3 : re_subsecc.CursorLocation = 2 : re_subsecc.LockType = 3
	re_subsecc.Open()
%>
                <table border="0" cellpadding="2" cellspacing="3">
                  <tr>
                    <%while not re_subsecc.eof
			if ""&seccion2 = ""&re_subsecc("S2_ID") then%>
                    <td class="boton-over"><b>&nbsp;<%=re_subsecc("S2_NOMBRE")%>&nbsp;</b></td>
                    <%else%>
                    <td class="boton-out"><a href="JavaScript:irSeccion(<%=seccion%>,<%=re_subsecc("S2_ID")%>)">&nbsp;<%=re_subsecc("S2_NOMBRE")%>&nbsp;</a></td>
                    <%end if%>
                    <%re_subsecc.movenext : wend
		if numero(seccion) = -1 then%>
                    <td class="boton-over">&nbsp;Todas&nbsp;</td>
                    <%else%>
                    <td class="boton-out">&nbsp;<a href="JavaScript:todasSecciones2()">Todas</a>&nbsp;</td>
                    <%end if%>
                  </tr>
                </table>
                <%
	re_subsecc.Close()
	set re_subsecc = Nothing
%>
            </td>
          </tr>
        </table></td>
      </tr>
    </table>
<%end if ' if seccion <> -1%>

<br>


<%if totalRegistros = 0 then%>
<center>
  <b>No hay ning&uacute;n resultado disponible</b><br>
  <br>
</center>
<%else
%>
      <table width="100%" border="0" cellpadding="2" cellspacing="0">
		<TR>
		<TD width="100%" colspan="2" class="general-pixel-abajo"><font color="#0099FF" size="1">Tema</font></TD>
		<TD class="general-pixel-abajo"><font color="#0099FF" size="1">Resp.</font></TD>
		<TD align="center" class="general-pixel-abajo"><font color="#0099FF" size="1">Fecha</font></TD>
		</tr>
		<tr><td height="5" colspan="6"><img src="../../spacer.gif" width="1" height="1"></td>
		</tr>
	<%for n=0 to registrosPorPagina-1
		if not re.eof then%>

		<tr>
		  
		  <td align="right" valign="top">
            <%if re("R_USUARIO") > 0 then%>
			  	<span title=" Código: <%=re("R_USUARIO")%> "><%=getNombreUsuario(re("R_USUARIO"))%></span>
			<%else%>
				An&oacute;nimo
			<%end if%>:</td>
		  <td width="100%" align="left" valign="top"><img src="/<%=c_s%>img/ico_carpeta.gif" width="18" height="15" border="0" align="absmiddle">
		    <%if nav_ampliar then%>
<a href="JavaScript:ampliar(<%=re("R_ID")%>);">
<%end if%>
<%if nav_ampliar then
			Response.Write unpoco(re("R_TITULO"),60)
		else
			Response.Write re("R_TITULO")
		end if%>
<%if nav_ampliar then Response.Write "</a>" end if%>
</a></td>
		  <td align="center" valign="top"><font color="#666666" size="1"><%=re("R_NUM_RESPUESTAS")%></font></td>
		  <td align="right" valign="top"><font color="#666666" size="1"><%=re("R_AUTOFECHA")%></font></td>
		</tr>


  <%
  re.movenext
			  
		end if
	next
	%>
</table>
<%end if%>

	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
	<tr>
                  <td align="right"><a href="index.asp?secc=<%=secc%>&ac=nuevo_tema&seccion=<%=seccion%>">Nuevo tema</a></td>
                </tr>
              </table>

<%if totalRegistros >  registrosPorPagina then%>
  <div align="center">Páginas: 
	<%
	' Numero de página -------------------------------------------------------------------------
	Dim salida, num ,n
	salida = ""
	num = 0
	for n=1 to totalRegistros
		if n mod registrosPorPagina = 1 then
			num = num +1
			if int(num) = int(pag) then%>
			  <b>[<%=num%>]</b>
		  <%else%>
			  <a href='#' onClick='irPag(<%=num%>)'><%=num%></a>
		  <%end if
		end if
	next
	' -------------------------------------------------------------------------------- Números de página%>
</div>
<%else
	'Response.Write("<div align='center'>Paginas: <b>[1]</b></div>")
end if

re.Close() : set re = Nothing
end select ' ------------------------------------------------------------------------------------ end select%>
                <!-- FIN DE FORMULARIO --------------------------------------------------------------------------------------- -->
<%if nav_busca then%>
	<script>f.cadena.focus()</script>
<%end if%>
<%end if ' unerror


sub pintaIconoExtension (ext)
		select case ext
			case "jpg","png","gif","bmp","tif"
				ico = "img"
			case "exe"
				ico = "jpg"
			case "doc"
				ico = "jpg"
			case "txt"
				ico = "jpg"
			case "xls"
				ico = "jpg"
			case "zip"
				ico = "zip"
			case "pdf"
				ico = "pdf"
			case "mp3","wav"
				ico = "mp3"
			case else
				ico = "iconootro"
		end select
			%><img src="/<%=c_s%>img/<%=ico%>.gif" align="absmiddle" alt=" Archivo tipo: <%=UCASE(ext)%> "><%
	end sub

end if

if unerror then%>
	<font color="#FF0000"><b>ATENCIÓN</b>:<br>
<%=msgerror%></font>
<%end if%>