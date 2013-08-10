<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' PARA EMPEZAR A FUNCIONAR DEBEMOS SABER LO QUE DESEAMOS ADMINISTRAR
Dim cualid ' Cualiadad, o zona a administrar. (EJ: Noticias, Archivos, ...)

if ""&request.QueryString("cualid")<>"" then
	cualid = request.QueryString("cualid")
	session("cualid") = cualid
elseif ""&session("cualid") <> "" then
	cualid = session("cualid")
else
	unerror = true : msgerror = "No se ha especificado una zona de aSkipper. (Cualidad)"
end if
%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include file="inc_inicia_xml.asp" -->
<%inicia_xml()%>
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<!--#include file="inc_seguridad.asp" -->
<!--#include virtual="/admin/inc_rutinas.asp" -->
<!--#include file="inc_rutinas.asp" -->
<!--#include file="inc_conn.asp" -->
<!--#include virtual="/admin/inc_sha256.asp" -->
<!--#include virtual="/admin/usuarios/inc_gestion_grupos.asp" -->

<%
Dim retotal
retotal = 0 ' registros encontrados en total
Dim xmlObj
Dim unerror, msgerror, errorform, msgerrorform
Dim sql, re, n, v, id, oConn, d, mfso, mfsoescribir
Dim titulo, subtitulo, cuerpo, seccion, fuente, orden, enlace, enportada, activo, duplicar, fecha, fechaini, fechafin, foto, icono
Dim ruta_xml_config ' RUTA ^
ruta_xml_config = Server.MapPath("/datos/xml_admin_config.xml")
Dim c_titulo, c_nombre, c_tipo, opcion, valor, valores, pos1, pos2, dire, msg_consulta, nombreSeccionActual
Dim fechactual
fechactual = date()
Dim arrMes
arrMes = array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")


Dim ac
ac = ""&request.Form("ac")
if ac = "" and ""&request.QueryString("ac") <> "" then
	ac = ""&request.QueryString("ac")
end if%>


<html>
<head>
<title>aSkipper</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="estilos.css" rel="stylesheet" type="text/css">
<script>

	//
	function ventana(dire,nombre,ancho,alto,barras) {
		var winl = (screen.width - ancho) / 2;
		var wint = (screen.height - alto) / 2;
		if (""+barras == "1"){
			barras = "yes"
		}else{
			barras = "no"
		}
		var paramet = "scrollbars="+ barras +",top="+ wint +",left="+ winl +",width="+ ancho +",height="+ alto;
		var splashWin = window.open(dire,nombre,paramet);
		splashWin.focus();
	}
	//
	function popMail(nombre, email, asunto, cuerpo, id){
		ventana("popMail.asp?nombre="+ nombre +"&email="+ email +"&asunto="+ asunto +"&cuerpo="+ cuerpo +"&id="+ id,"popMail",500,320,0)
	}

// Obtener las posiciones X,Y de un objeto relativo (Marc Palau)
function getPos(id){
   var o=document.getElementById(id);
   var oLeft=o.offsetLeft;
   var oTop=o.offsetTop;
   while(o.offsetParent.style.position.toLowerCase()!="absolute"){
      oParent=o.offsetParent;
      oLeft+=oParent.offsetLeft;
      oTop+=oParent.offsetTop;
      o=oParent;
   };
   return [oLeft,oTop];
};


// -------------------------------------------------------------------------------  AREA HTML (EDITOR)
_editor_url = "";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
	document.write('<scr' + 'ipt src="' +_editor_url+ 'inc_editorHtml.js"');
	document.write(' language="Javascript1.2"></scr' + 'ipt>');
} else {
	document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>');
}

misBotones = [
	//	['fontname'],
	//	['fontsize'],
	//	['fontstyle'],
	//  ['linebreak'],
	['bold','italic','underline','separator'],
	['Createlink'],
	//	['mayusminus','separator'],
	//	['strikethrough','subscript','superscript','separator'],
	['justifyleft','justifycenter','justifyright','separator'],
    //['OrderedList','UnOrderedList','Outdent','Indent','separator'],
    ['forecolor','backcolor','separator'],
    ['HorizontalRule',/*,'InsertImage','InsertTable'*/],
   // ['InsertImage'],
    ['insertaskimage'],
//  ['custom1','custom2','custom3','separator'],
	['popupeditor'],
    ['intro'],
//    ['about'],
    ['htmlmode']
	];

//-------------------------------------------------------------------------------------------------------- -->

</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<style type="text/css">
<!--
.Estilo8 {color: #006600}
.Estilo9 {color: #660000}
.TituloResaltado {
	font-size: 11pt;
	font-weight: bold;
	color: #849ACE;
}

-->
</style>
</head>
<body class="bodyAdmin" vlink="#003366">



<%if unerror then
	Response.Write "<b>Error</b><br>"&msgerror
else
	
	' VARIABLES Y RECUPERACIÓN
	'----------------------------------------------------------------------------------------
	
	' Seccion actual
	Dim idSeccionActual
	idSeccionActual = numero(request.Form("seccion"))
	
	' Seccion2 actual
	Dim idSeccion2Actual
	idSeccion2Actual = 	numero(request.Form("seccion2"))
	
	' Cadana del campo buscar

	
	Dim cadena, cadenabak
	cadena = ""&replace(trim(request.Form("cadena")),"'","") ' quito las comillas simples

	' manolo
	cadenabak = cadena
	cadena = replace(cadena,chr(34),"")' quito las comillas dobles

	
	' Registros por página
	Dim regporpag
	regporpag = config_regporpag

	if isNumeric(request.Form("regporpag")) and not ""&request.Form("regporpag") = "" then
		regporpag = int(request.Form("regporpag"))
	end if
	
	' Inicio de registro 
	Dim regInicio
	regInicio = 1
	if isNumeric(request.Form("reginicio")) and not ""&request.Form("reginicio") = "" then
		regInicio = int(request.Form("reginicio"))
	end if
	
	'----------------------------------------------------------------------------------------

%>
	<!-- Todo por formularios -->
	<form action="#" method="post" name="f" id="f" onSubmit="return envio()">
	<input name="ac" id="ac" type="hidden" value="">
	<input type="hidden" name="msgerror" value="">
	<input type="hidden" name="id" value="">
  <input name="reginicio" type="hidden" id="reginicio" value="<%=regInicio%>">
  <input name="cadenabak" type="hidden" id="cadenabak" value="<%=cadenaBak%>">
  <input type="hidden" name="portada" value="<%=request.Form("portada")%>">
	<input type="hidden" name="pag" value="<%=request.Form("pag")%>">
	<input type="hidden" name="comun" value="">
	<input type="hidden" name="comun2" value="">
<%select case ac

	case "duplicar"%>
		<!--#include file="inc_main_duplicar.asp" -->

	<%case "borrarseleccion"%>
		<!--#include file="inc_main_borrarseleccion.asp" -->

	<%case "eliminar"%>
		<!--#include file="inc_main_eliminar.asp" -->

	<%case "editar"%>
		<!--#include file="inc_main_editar.asp" -->
	
	<%case "nuevo"%>
		<!--#include file="inc_main_nuevo.asp" -->

	<%case "ampliar"%>
		<!--#include file="inc_main_ampliar.asp" -->

	<%case "adminordenidioma"%>
		<!--#include file="inc_main_adminordenidioma.asp" -->

	<%case "adminsecciones"%>
		<!--#include file="inc_main_adminsecciones.asp" -->
	
	<%case "adminsecciones2"%>
		<!--#include file="inc_main_adminsecciones2.asp" -->

	<%case "ordenidiomavisible"
	
		call exeSql("UPDATE ORDENIDIOMA SET OI_VISIBLE_"&Ucase(request.Form("idi"))&" = "& request.Form("v") &" WHERE OI_ID = "& request.Form("id"), conn_)%>
		<script language="javascript" type="text/javascript">
			f.ac.value = "adminordenidioma"
			f.submit()
		</script>

	<%case "eliminarordenidioma"

		id = request.QueryString("id")
		if id <> "" and isNumeric(id) then
			' Borro
			call exeSql("DELETE FROM ORDENIDIOMA WHERE OI_ID = "& id, conn_)
			reordenaOrdenIdioma()
			%>
			<script>
			try{
				var f1 = parent.frames[0].f // Frame de la izquierda
				f1.ac.value = ""
				f1.action = "main.asp"
				f1.target = ""
				f1.submit()
			}catch(unerror){}
	
			f.ac.value = "adminordenidioma"
			f.submit()
			</script>
			<%
		end if

	case "eliminarseccion"

		id = request.QueryString("id")
		if id <> "" and isNumeric(id) then
			' Compruebo si hay registros vinculados a esta sección
			consultaXOpen "SELECT R_SECCION FROM REGISTROS WHERE R_SECCION = "& id,1
			if not re.eof then
				consultaXClose()
				%>
				<script>
				f.ac.value = "adminsecciones"
				f.msgerror.value = "Hay registros vinculados a esta sección."
				f.submit()
				</script>
				<%
			else
				consultaXClose()
				
				' Borro
				call exeSql("DELETE FROM SECCIONES WHERE S_ID = "& id, conn_)
				reordenaSecciones()
				%>
				<script>
				try{
					var f1 = parent.frames[0].f // Frame de la izquierda
					f1.ac.value = ""
					f1.action = "main.asp"
					f1.target = ""
					f1.submit()
				}catch(unerror){}
	
				f.ac.value = "adminsecciones"
				f.submit()
				</script>
				<%
			end if		
		end if
	
	case "eliminarseccion2"

		seccion = numero(request("seccion"))
		id = request.QueryString("id")
		if esNumero(id) and seccion >0 then
			' Compruebo si hay registros vinculados a esta sección
			consultaXOpen "SELECT R_SECCION2 FROM REGISTROS WHERE R_SECCION = "& seccion &" AND R_SECCION2 = "& id,1
			if not re.eof then
				consultaXClose()
				Response.Redirect("main.asp?ac=adminsecciones2&seccion="& seccion &"&msgerror=Hay registros vinculados a esta sub-sección.<br>Debe borrarlos antes.")
			else
				consultaXClose()
				
				' Borro
				call exeSql("DELETE FROM SECCIONES2 WHERE S2_ID = "& id, conn_)
				reordenaSecciones2(seccion)
				downSeccion(seccion)
				%>
				<script language="javascript" type="text/javascript">
				try{
					var f1 = parent.frames[0].f // Frame de la izquierda
					f1.ac.value = ""
					f1.action = "main.asp"
					f1.target = ""
					f1.submit()
				}catch(unerror){}
	
				location.href='main.asp?ac=adminsecciones2&seccion=<%=seccion%>'
				</script>
				<%
			end if		
		end if


	case "editarordenidioma"%>
		<!--#include file="inc_main_editarordenidioma.asp" -->

	<%case "editarseccion"%>
		<!--#include file="inc_main_editarseccion.asp" -->

	<%case "editarseccion2"%>
		<!--#include file="inc_main_editarseccion2.asp" -->

	<%case "moverordenidioma"


		id = request.Form("id")
		dire = ""&request.Form("comun")
		idiactual = ucase(""&request.Form("idiactual"))
		if id <> "" and isNumeric(id) and dire <> "" then
		
			' Cojo todas las secciones
			consultaXOpen "SELECT OI_ID, OI_ORDEN_ESP, OI_ORDEN_ENG, OI_ORDEN_FRA, OI_ORDEN_DEU, OI_ORDEN_ITA, OI_BLOQUEADA FROM ORDENIDIOMA ORDER BY OI_ORDEN_"& idiactual,2
			
			while not re.eof
				' Si estoy en el registro solicitado
				if re("OI_ID") = int(id) then
					' Para subir
					if dire = "subir" then
						' Si no estoy en primer lugar
						if re("OI_ORDEN_"& idiactual) > 1 then
							
							' Declaro el estado de Bloqueo de la sección anterior
							pasos = 1
							v = true ' Busca
							re.movePrevious()					
							while not re.bof and v
								if re("OI_BLOQUEADA") = true then
									pasos = pasos + 1
									re.movePrevious()
								else
									v = false
								end if
							wend
		
							' Volvemos a la posición de nuestra sección
							pasos = pasos - 1 ' (Por el primer movePrevious())
							for n=0 to pasos
								re.moveNext()
							next
		
							' Nos movemos a el lugar indicado saltando las secciones bloqueadas
							re("OI_ORDEN_"& idiactual) = re("OI_ORDEN_"& idiactual) - (1.5 * (pasos+1))
							re.update()
							valor = true
						end if
					else
							Response.Write "<br>reTotal " & re("OI_ORDEN_"& idiactual) & " " & reTotal
						' Si no estoy el último
						if re("OI_ORDEN_"& idiactual) < reTotal then
							Response.Write "<br>No soy la última"
							
							' Declaro el estado de Bloqueo de la sección anterior
							pasos = 1
							v = true ' Busca
							re.moveNext()					
							while not re.eof and v
								Response.Write "<br>He dado un paso"
								if re("OI_BLOQUEADA") = true then
									Response.Write "<br>He encontrado una bloq"
									pasos = pasos + 1
									re.moveNext()
								else
									Response.Write "<br>Paro de buscar"
									v = false
								end if
							wend
		
							' Volvemos a la posición de nuestra sección
							pasos = pasos - 1 ' (Por el primer moveNext())
							for n=0 to pasos
								re.movePrevious()
							next
							Response.Write "<br>Pasos "& pasos
							' Nos movemos a el lugar indicado saltando las secciones bloqueadas
							re("OI_ORDEN_"& idiactual) = re("OI_ORDEN_"& idiactual) + (1.5 * (pasos+1))
							re.update()
							valor = true
						end if
					end if
				end if
				re.movenext
			wend
			reordenaOrdenIdioma()
		end if
		
		%>
		Un momento ...
		<input type="hidden" name="idiactual" value="<%=request.Form("idiactual")%>">
		<script language="javascript" type="text/javascript">
		<!--
			f.ac.value = "adminordenidioma"
			f.submit()
		//-->
		</script>


<%case "moverseccion"


id = request.Form("id")
dire = ""&request.Form("comun")
if id <> "" and isNumeric(id) and dire <> "" then

	' Cojo todas las secciones
	consultaXOpen "SELECT S_ID, S_ORDEN, S_BLOQUEADA FROM SECCIONES ORDER BY S_ORDEN",2
	
	while not re.eof
		' Si estoy en el registro solicitado
		if re("S_ID") = int(id) then
			' Para subir
			if dire = "subir" then
				' Si no estoy en primer lugar
				if re("S_ORDEN") > 1 then
					
					' Declaro el estado de Bloqueo de la sección anterior
					pasos = 1
					v = true ' Busca
					re.movePrevious()					
					while not re.bof and v
						if re("S_BLOQUEADA") = true then
							pasos = pasos + 1
							re.movePrevious()
						else
							v = false
						end if
					wend

					' Volvemos a la posición de nuestra sección
					pasos = pasos - 1 ' (Por el primer movePrevious())
					for n=0 to pasos
						re.moveNext()
					next

					' Nos movemos a el lugar indicado saltando las secciones bloqueadas
					re("S_ORDEN") = re("S_ORDEN") - (1.5 * (pasos+1))
					re.update()
					valor = true
				end if
			else
					Response.Write "<br>reTotal " & re("S_ORDEN") & " " & reTotal
				' Si no estoy el último
				if re("S_ORDEN") < reTotal then
					Response.Write "<br>No soy la última"
					
					' Declaro el estado de Bloqueo de la sección anterior
					pasos = 1
					v = true ' Busca
					re.moveNext()					
					while not re.eof and v
						Response.Write "<br>He dado un paso"
						if re("S_BLOQUEADA") = true then
							Response.Write "<br>He encontrado una bloq"
							pasos = pasos + 1
							re.moveNext()
						else
							Response.Write "<br>Paro de buscar"
							v = false
						end if
					wend

					' Volvemos a la posición de nuestra sección
					pasos = pasos - 1 ' (Por el primer moveNext())
					for n=0 to pasos
						re.movePrevious()
					next
					Response.Write "<br>Pasos "& pasos
					' Nos movemos a el lugar indicado saltando las secciones bloqueadas
					re("S_ORDEN") = re("S_ORDEN") + (1.5 * (pasos+1))
					re.update()
					valor = true
				end if
			end if
		end if
		re.movenext
	wend
	reordenaSecciones()
end if

	Response.Redirect("main.asp?ac=adminsecciones")
	
case "moverseccion2"

seccion = numero(request("seccion"))
id = request.Form("id")
dire = ""&request.Form("comun")
if id <> "" and isNumeric(id) and dire <> "" then

	' Cojo todas las secciones
	consultaXOpen "SELECT S2_ID, S2_ORDEN, S2_BLOQUEADA FROM SECCIONES2 ORDER BY S2_ORDEN",2
	
	while not re.eof
		' Si estoy en el registro solicitado
		if re("S2_ID") = int(id) then
			' Para subir
			if dire = "subir" then
				' Si no estoy en primer lugar
				if re("S2_ORDEN") > 1 then
					
					' Declaro el estado de Bloqueo de la sección anterior
					pasos = 1
					v = true ' Busca
					re.movePrevious()					
					while not re.bof and v
						if re("S2_BLOQUEADA") = true then
							pasos = pasos + 1
							re.movePrevious()
						else
							v = false
						end if
					wend

					' Volvemos a la posición de nuestra sección
					pasos = pasos - 1 ' (Por el primer movePrevious())
					for n=0 to pasos
						re.moveNext()
					next

					' Nos movemos a el lugar indicado saltando las secciones bloqueadas
					re("S2_ORDEN") = re("S2_ORDEN") - (1.5 * (pasos+1))
					re.update()
					valor = true
				end if
			else
					'Response.Write "<br>reTotal " & re("S2_ORDEN") & " " & reTotal
				' Si no estoy el último
				if re("S2_ORDEN") < reTotal then
					Response.Write "<br>No soy la última"
					
					' Declaro el estado de Bloqueo de la sección anterior
					pasos = 1
					v = true ' Busca
					re.moveNext()					
					while not re.eof and v
						Response.Write "<br>He dado un paso"
						if re("S2_BLOQUEADA") = true then
							Response.Write "<br>He encontrado una bloq"
							pasos = pasos + 1
							re.moveNext()
						else
							Response.Write "<br>Paro de buscar"
							v = false
						end if
					wend

					' Volvemos a la posición de nuestra sección
					pasos = pasos - 1 ' (Por el primer moveNext())
					for n=0 to pasos
						re.movePrevious()
					next
					Response.Write "<br>Pasos "& pasos
					' Nos movemos a el lugar indicado saltando las secciones bloqueadas
					re("S2_ORDEN") = re("S2_ORDEN") + (1.5 * (pasos+1))
					re.update()
					valor = true
				end if
			end if
		end if
		re.movenext
	wend
	reordenaSecciones2(seccion)
end if

	Response.Redirect("main.asp?ac=adminsecciones2&seccion="&seccion)

case "moverregistro"

id = numero(request.Form("id"))
dire = ""&request.Form("comun")

mi_seccion = 0
mi_seccion2 = 0

if esNumero(id) and dire <> "" then

	' Se ordenará según la sección actual.
	if idSeccion2Actual > 0 then
		campo = "R_ORDEN_SECCION2"
	elseif idSeccionActual > 0 then
		campo = "R_ORDEN_SECCION"
	else
		campo = "R_ORDEN"
	end if

	consultaXOpen "SELECT R_ID, "& campo &", R_SECCION, R_SECCION2 FROM REGISTROS WHERE R_ID = "& id &" ORDER BY "& campo &"",2

	if not re.eof then
		mi_seccion = re("R_SECCION")
		mi_seccion2 = re("R_SECCION2")
		n = re(campo)
		if dire = "subir" then
			Response.Write "-"
			re(campo) = n - 1.5
		else
			re(campo) = n + 1.5
		end if
		re.update()
		valor = true
	end if
%>

Cambiando posición.<br>
Un momento ...

<script>
	try{
		var f = top.frames[1].frames[0].f // Frame de la izquierda
		f.ac.value = ""
		f.action = "main.asp"
		f.target = ""
		f.submit()
		location.href = 'inicio.asp'
	}catch(unerror){}
</script>
<%
	consultaXClose()
	' Reorganiza el campo orden
	campo_orden = campo
	reOrdena()

end if

case "iorden"%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
		<td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Intercambiar
			  orden</font></b></td>
		<td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
	  </tr>
	</table>
	<br>

<%
pos1 = 0 + ("0"&request.QueryString("pos1"))
pos2 = 0 + ("0"&request.QueryString("pos2"))

if pos1 > 0 and pos2 > 0 then
	
	consultaXOpen "SELECT * FROM REGISTROS WHERE R_ORDEN = "& pos1 &" OR R_ORDEN = "& pos2 &" ORDER BY R_ORDEN",2
	
	if retotal = 2 then
		if re("R_ORDEN") = pos1 then
			re("R_ORDEN") = pos2
			re.movenext
			re("R_ORDEN") = pos1
		else
			re("R_ORDEN") = pos1
			re.movenext
			re("R_ORDEN") = pos2
		end if
		re.update()
	end if
	consultaXClose()
	reOrdena()
	if err<>0 then
		%>
		<b>Error</b>.<br>
		No se ha podido realizar el cambio.
		<%
	else
	%>
	<br>
	<br>
	<div align="center"><font color="#006600"><b>Cambio realizado</b></font></div>
	<script>
	function cerrar(){
		try{
			var f = parent.opener.f
			f.target = ""
			f.submit()
		}catch(unerror){
			//
		}
		window.close()
	}
	setTimeout("cerrar()",500)
	</script>
	<%
	end if
	
else
	consultaXOpen "SELECT R_ORDEN FROM REGISTROS ORDER BY R_ORDEN",1
	v = request.QueryString("r")
	%>
	<br>
	<table border="0" align="center" cellpadding="4" cellspacing="0">
	  <tr>
		<td>Cambiar el registro</td>
		<td><select name="pos1" class="campoAdmin" id="pos1">
		<%for n=1 to retotal%>
		  <option value="<%=n%>" <%if ""&v = ""&n then Response.Write "selected" end if%> ><%=n%></option>
		  <%next%>
			</select></td>
	  </tr>
	  <tr>
		<td>Por el registro </td>
		<td>
		<select name="pos2" class="campoAdmin" id="pos2">
		<%for n=1 to retotal%>
			<option value="<%=n%>"><%=n%></option>
		<%next%>
		</select></td>
	  </tr>
	</table>
    <br>	
    <br>
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
	  <tr>
		<td align="right">		  <input name="" type="button" class="botonAdmin" onClick="iorden_ya(f.pos1.value,f.pos2.value)" value="Aceptar"></td></tr>
	</table>
	<br>
	<%
	consultaXClose()
end if

case "verenlacefoto"

	foto = request.QueryString("foto")
	host = "http://"& request.ServerVariables("HTTP_HOST")

	v = "/" & c_s & "descargas/?idi="&session("idioma")&"&cualid="&cualid&"&id="& id
	v = "/" & c_s & "datos/"&session("idioma")&"/"&cualid&"/fotos/"&foto
	enlaceexterno =  host & v
	enlaceinterno = v
%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
        <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Enlazar imagen </font></b></td>
        <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
      </tr>
    </table>

	<br>
	<table width="97%" border="0" align="center" cellpadding="1" cellspacing="0">
      <tr>
        <td colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2">Enlace</td>
      </tr>
      <tr>
        <td><input name="textfield" type="text" class="campoAdmin" value="<%=enlaceexterno%>" size="75">
        </td>
        
      </tr>
    </table>
	<br>

<%case "verenlacearchivo"

	id = request.QueryString("id")
	host = request.ServerVariables("HTTP_HOST")
	enlaceexterno =  "http://"& host & "/" & c_s & "descargas/?idi="&session("idioma")&"&cualid="&cualid&"&id="& id
%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
        <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Enlace de descarga</font></b></td>
        <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
      </tr>
    </table>

	<br>
	<table width="97%" border="0" align="center" cellpadding="1" cellspacing="0">
      <tr>
        <td colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2">Enlace de descarga</td>
      </tr>
      <tr>
        <td><input name="textfield" type="text" class="campoAdmin" value="<%=enlaceexterno%>" size="60">
        </td>
        <td align="right"><nobr>&nbsp;<a href="<%=enlaceexterno%>"><img src="img/descargar.gif" width="79" height="18" border="0" align="absbottom"></a></nobr></td>
      </tr>
    </table>
	<br>


<%case "verenlace"

	id = request.QueryString("id")
	host = request.ServerVariables("HTTP_HOST")
	host = "http://"& host & "/" & c_s
	v = "index.asp?secc=/"& cualid &"&id=" & id
	enlaceexterno =  host & "/" & session("idioma")& "/"& v
	enlaceinterno = v
%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="8" height="19"><img src="img/titulo_izq.gif" width="8" height="19"></td>
        <td align="center" valign="middle" background="img/titulo_cen.gif"><b><font color="#FFFFFF">Enlace a registro</font></b></td>
        <td width="8" height="19"><img src="img/titulo_der.gif" width="8" height="19"></td>
      </tr>
    </table>

	<br>
	<table width="95%" border="0" align="center" cellpadding="1" cellspacing="0">
      <tr>
        <td colspan="3">Enlace externo</td>
      </tr>
      <tr>
        <td><input name="textfield" type="text" style="width=100%" value="<%=enlaceexterno%>" size="70">
        </td>
        <td align="right" bgcolor="#FFFFFF"><a href="<%=enlaceexterno%>" target="_blank" class="aAdmin">Ir <img src="img/flecha_der.gif" width="16" height="16" border="0" align="absmiddle"></a></td>
        <td width="5" align="right" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3">Enlace interno</td>
      </tr>
      <tr>
        <td><input name="textfield" type="text" style="width=100%" value="<%=enlaceinterno%>" size="70"></td>
        <td align="right" bgcolor="#FFFFFF"><a href="../../<%=session("idioma")%>/<%=enlaceinterno%>" target="_blank" class="aAdmin">Ir <img src="img/flecha_der.gif" width="16" height="16" border="0" align="absmiddle"></a></td>
        <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
    </table>
	<br>


<%case else

	' Cargamos las secciones a un XML
	if not unerror then

		sql = "SELECT * FROM SECCIONES ORDER BY S_ORDEN"
		set re = Server.CreateObject("ADODB.Recordset")
		on error resume next
		re.ActiveConnection = conn_
		if err<> 0 then
			unerror = true : msgerror = "Error en conexión a base de datos.<br>"&err.description&".<br>"&conn_
		else
			re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 1
			re.Open()
			if err <> 0 then
				unerror = true : msgerror = "Error en SQL.<br>"&err.description&".<br>"&sql
			end if
		end if
		on error goto 0
		
		if not unerror then
			Dim totalSecciones : totalSecciones = 0
			retotal = re.recordcount
			if retotal > 0 then
			totalSecciones = retotal
				' Creo un XML con las secciones
				dim xmlSecciones, padreSecciones, nodoSeccion, attNombre
				set xmlSecciones = CreateObject("MSXML.DOMDocument")
				set padreSecciones = xmlSecciones.createElement("secciones")
	
				for n=1 to totalSecciones
					set nodoSeccion = xmlSecciones.createElement("id"& re("S_ID"))
					padreSecciones.appendChild(nodoSeccion)
					set attNombre = xmlSecciones.createAttribute("nombre")
					nodoSeccion.setAttributeNode(attNombre)
					attNombre.nodeValue = re("S_NOMBRE")
					set attNombre = nothing
					
					set attNuevos = xmlSecciones.createAttribute("nuevos")
					nodoSeccion.setAttributeNode(attNuevos)
					attNuevos.nodeValue = re("S_NUEVOS")
					set attNombre = nothing
	
					set nodoSeccion = nothing
					re.moveNext()
				next
				xmlSecciones.appendChild(padreSecciones)
	'			xmlSecciones.save Server.MapPath("prueba.xml")
			end if
			re.close()
			set re = nothing
		end if ' unerror 

		if unerror then
'			msgerror = "Error aki"
		end if
	end if
	
		function getIdSeccion(id)
			getIdSeccion = numero(right(id,len(id)-2))
		end function
		
		function nuevosSeccion(id)	
			dim nuevos, nodo
			nuevos = false
			if ""&id <> "" then
				if typeOK(xmlSecciones) then
					set nodo = xmlSecciones.selectSingleNode("/secciones/id"&id)
					if typeOK(nodo) then
						nuevos =  cbool(numero(""&nodo.getAttribute("nuevos")))
					end if
				end if
			end if
			nuevosSeccion = nuevos
		end function

	
	if not unerror then
	
		retotal = 0
	
		sql = "SELECT *"
		sql = sql & " FROM REGISTROS, SECCIONES"
		sql = sql & " WHERE (R_SECCION = S_ID)"
	
		' Sección
		if idSeccionActual > 0 then
			sql = sql & " AND R_SECCION = "& idSeccionActual
		end if
	
		' Sección2
		if idSeccion2Actual > 0 then
			sql = sql & " AND R_SECCION2 = "& idSeccion2Actual
		end if
	
		' Cadena de búsqueda
		if cadena <> "" then
			if config_cuerpo <> "" then
				sql = sql & " AND (R_TITULO LIKE '%"& cadena &"%' OR R_"& config_cuerpo &" LIKE '%"& cadena &"%')"
			else
				sql = sql & " AND R_TITULO LIKE '%"& cadena &"%'"
			end if
		end if
		
		if ""&request.Form("portada") = "1" then
			sql = sql & " AND R_PORTADA = 1"
		end if
		
		if config_estados then
			if ""& request.Form("estado") <> "" then
				sql = sql & " AND R_ESTADO = '"& request.Form("estado") &"'"
			end if
		end if
	
		' sql extra
		if config_sqlextra <> "" then
			sql = sql & config_sqlextra
		end if

		' Si sólo tememos una sección, la tomo como actual
		if typeOK(padreSecciones) then
			if padreSecciones.childNodes.length = 1 then
'				idSeccionActual = getIdSeccion(padreSecciones.childNodes.item(0).nodeName)
				launica = getIdSeccion(padreSecciones.childNodes.item(0).nodeName)
			end if
		end if
	
		' Orden
		if config_sqlOrden <> "" then
			sql = sql & " ORDER BY " & config_sqlorden
		elseif idSeccion2Actual > 0 then
			sql = sql & " ORDER BY R_ORDEN_SECCION2"
		elseif idSeccionActual > 0 then
			sql = sql & " ORDER BY R_ORDEN_SECCION"
		else
			sql = sql & " ORDER BY R_ORDEN"
		end if
	
		
		' Leemos las subSecciones (secciones2)
		' Necesito este dato: idSeccionActual
		if config_activo_seccion2 then
			' Cargamos las secciones2 a un XML
			if not unerror then
				consultaXOpen "SELECT * FROM SECCIONES2 WHERE S2_ID_S = "& idSeccionActual &" ORDER BY S2_ORDEN",1
				Dim totalSecciones2 : totalSecciones2 = 0
				if retotal > 0 then
				totalSecciones2 = retotal
					' Creo un XML con las secciones
					dim xmlSecciones2, padreSecciones2, nodoSeccion2, attNombre2
					set xmlSecciones2 = CreateObject("MSXML.DOMDocument")
					set padreSecciones2 = xmlSecciones2.createElement("secciones")
		
					for n=1 to totalSecciones2
						set nodoSeccion2 = xmlSecciones2.createElement("id"& re("S2_ID"))
						padreSecciones2.appendChild(nodoSeccion2)
						set attNombre2 = xmlSecciones2.createAttribute("nombre")
						nodoSeccion2.setAttributeNode(attNombre2)
						attNombre2.nodeValue = re("S2_NOMBRE")
						set attNombre2 = nothing
						set nodoSeccion2 = nothing
						re.moveNext()
					next
					xmlSecciones2.appendChild(padreSecciones2)
		'			xmlSecciones.save Server.MapPath("prueba.xml")
				end if
				consultaXClose()
				
				function getIdSeccion2(id)
					getIdSeccion2 = numero(right(id,len(id)-2))
				end function
			end if
		end if
		
		alt_sql = sql
		
		consultaXOpen sql,1
		if unerror then
			Response.Write msgerror
		else
			' loop de pintado
			Dim regLoop
			' Si el total de registros encontrado es menor que el paginado seleccionado.
			if (retotal - (regInicio-1)) < regporpag then
				regLoop = retotal - (regInicio-1)
			else
				regLoop = regporpag
			end if
	
			if regInicio <= retotal then
				re.move regInicio - 1
			end if	



		end if	' unerror
	%>

<table width="100%" height="18"  border="0" cellpadding="0" cellspacing="0" class=fondoOscuroCblancoAdmin>
  <tr>
    <td width="4"><img src="img/spacer.gif" width="4" height="3"></td>
    <td width="33%"><%=config_nombresitio%></td>
    <td width="33%" align="center"><b><%=config_nombrecualid%></b>
	<%
	if config_idioma_bd <> "" then
		Response.Write " <span title='Idioma fuente de la Base de Datos'>("& config_idioma_bd &")</span>"
	end if
	%></td>
    <td width="33%" align="right"><%=getNombreIdioma(session("idioma"))%></td>
    <td width="4" align="right"><img src="img/spacer.gif" width="4" height="8"></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<td><img src="../global/img/spacer.gif" width="1" height="4"></td>
	</tr>
</table>

<%if totalsecciones >0 then%>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="left"><table  border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td>
		<select name="seccion" class="campoAdmin" id="seccion" onChange="changeSeccion(this)" <%if not totalsecciones >0 then%>disabled="disabled"<%end if%>>
			<%'if totalsecciones >1 then%>
				<option value=""><%=config_nom_secciones%> ...</option>
			<%'end if%>
			<%dim a : for each a in padreSecciones.childNodes
				ids = getIdSeccion(a.nodeName)%>
				<option value="<%=ids%>" <%if getIdSeccion(a.nodeName) = idSeccionActual then Response.Write "selected" : nombreSeccionActual = a.getAttribute("nombre") end if%>><%=a.getAttribute("nombre")%></option>
			<%next%>
		</select>
          </td>
		  
		<%if totalSecciones2 > 0 then%>
          <td>
			<select name="seccion2" class="campoAdmin" id="seccion2" onChange="changeSeccion2(this)" <%if not totalsecciones >0 then%>disabled="disabled"<%end if%>>
	            <option value=""><%=config_nom_secciones%> ...</option>
	            <%for each a in padreSecciones2.childNodes%>
	            <option value="<%=getIdSeccion2(a.nodeName)%>" <%if getIdSeccion2(a.nodeName) = idSeccion2Actual then Response.Write "selected" : nombreSeccion2Actual = a.getAttribute("nombre") end if%>><%=a.getAttribute("nombre")%></option>
	            <%next%>
			</select>
          </td>
		<%end if%>

          <td><table border="0" cellpadding="2" cellspacing="0" class="botonAdmin" title=" Anula la búsqueda actual y muestra los registros de todas las secciones ">
            <tr>
              <td><a href="#" class="blanco" onClick="verTodo();return false;">Todo</a></td>
            </tr>
          </table></td>

        </tr>
      </table></td> 
      <td align="right">
		<%if config_buscar then%>
			<input name="cadena" type="text" class="campoAdmin" id="cadena" accesskey="b" title=" Escriba la palabra o palabras que desea buscar " value="<%=cadena%>" size="12" maxlength="30">
			<input type="submit" class="botonAdmin" value="Buscar">
			<%if cadena <> "" then%>
				<input name="" type="button" class="botonAdmin" title=" Cancelar la búsqueda actual " onClick="borrarCadena()" value="X"> 
			<%end if%>
		<%else%>
			<input name="cadena" type="hidden" id="cadena" value="<%=cadena%>">
		<%end if%>
	  </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><img src="../global/img/spacer.gif" width="1" height="2"></td>
    </tr>
  </table>
	  <%if config_buscar then%>
	  	<script>f.cadena.focus()</script>
	 <%end if%>
  

  <%else%>
  <input name="cadena" type="hidden" id="cadena" value="<%=cadena%>">
 <%end if ' if totalsecciones >0 then%>
  

<%if totalsecciones <= 0 then%>
  <table width="100%" border="0" cellpadding="1" cellspacing="0" bgcolor="#CC0000">
    <tr>
      <td><table width="100%"  border="0" cellspacing="0" cellpadding="2">
          <tr>
            <td bgcolor="#FFFFFF">Para empezar a crear nuevos registros debe insertar al menos una sección. Pulse en el bot&oacute;n <b>Secciones</b>. </td>
          </tr>
      </table></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img src="../global/img/spacer.gif" width="1" height="4"></td>
    </tr>
  </table>
  <%end if%>

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>
	  <table border="0" cellspacing="0" cellpadding="1">
        <tr>
		
	<%if totalsecciones >0 and config_nuevos and nuevosSeccion(idSeccionActual) = true then%>
		<td>
		<table border="0" cellpadding="2" cellspacing="0" class="botonAdmin" title=" Crear nuevo registro ">
			<tr>
			<td><a href="#" class="blanco" onClick="nuevo();return false;">Nuevo</a></td>
			</tr>
		</table>
		</td>
	<%end if%>

          <%if config_activo_seccion then%>
          <td>
            <table border="0" cellpadding="2" cellspacing="0" class="botonAdmin">
              <tr>
                <td><a href="#" class="blanco" onClick="adminSecciones();return false;"><%=config_nom_secciones%></a></td>
              </tr>
          </table></td>
          <%end if%>

          <%if config_activo_seccion2 then%>
          <td>
            <table border="0" cellpadding="2" cellspacing="0" class="botonAdmin">
              <tr>
                <td><a href="#" class="blanco" onClick="adminSecciones2();return false;"><%=config_nom_secciones2%></a></td>
              </tr>
          </table></td>
          <%end if%>

          <%if config_ordenidioma then%>
          <td>
            <table border="0" cellpadding="2" cellspacing="0" class="botonAdmin" title=" Administrar secciones ">
              <tr>
                <td><a href="#" class="blanco" onClick="adminOrdenIdioma();return false;"><nobr>Orden idioma</nobr></a></td>
              </tr>
          </table></td>
          <%end if%>

		<td>
		<%if totalsecciones >0 and config_nuevos then
			if config_alfinal <> "" then%>
				<input type="hidden" name="nuevosalfinal" value="<%=config_alfinal%>">
			<%else%>
				<table width="100%"  border="0" cellspacing="0" cellpadding="0" title=" Insertar registros al final ">
					<tr>
					<td><input name="nuevosalfinal" type="checkbox" id="nuevosalfinal" onClick="changeAlfinal(this)" value="1" <%if ""&request.Form("nuevosalfinal") = "1" then%>checked<%end if%>></td>
					<td><label for="nuevosalfinal">Al final</label></td>
					</tr>
				</table>
			<%end if
		end if%>
		</td>

        </tr>
      </table></td>
	  
<%if totalsecciones >0 then%>
      <td align="right">
        <table  border="0" cellspacing="0" cellpadding="1" title=" Número de registros que se muestran en cada página ">
          <tr>
            <td>Mostrar:</td>
            <td><select name="regporpag" class="campoAdmin" id="regporpag" onChange="changeRegPorPag(this)">
              <%for n=0 to config_regporpag_opt-1
	  	v = v +config_regporpag %>
              <option value="<%=v%>" <%if v = regporpag then Response.Write("selected") end if%>><%=v%></option>
              <%next%>
            </select></td>
          </tr>
        </table></td>
          <%end if%>
		
    </tr>
  </table>
  <table width="100%" height="1" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td class="fondoOscuroAdmin"><img src="../global/img/spacer.gif" width="1" height="1"></td>
    </tr>
  </table>
<%
if config_estados then
	if typeOK(nodoConfig) then
		set nodoEstados = nodoConfig.selectNodes("//dato[@campo='estado']").item(0)
		if typeOK(nodoEstados) then%>

			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				<td><img src="../global/img/spacer.gif" width="1" height="4"></td>
				</tr>
			</table>

			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				<td align="right">
				<span title="<%=request.Form("estado")%>">Filtrar por estado:</span>
				<select name="estado" class="campoAdmin" onChange="changeEstado(this)">
					<option value="">Seleccione uno...</option>
				<%for each nodo in nodoEstados.childNOdes%>
					<option value="<%=nodo.getAttribute("valor")%>" <%if ""& request.Form("estado") = ""& nodo.getAttribute("valor") then Response.Write "selected" end if%>><%=nodo.getAttribute("titulo")%></option>
				<%next%>
				</select>
				</td>
				</tr>
			</table>

			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				<td><img src="../global/img/spacer.gif" width="1" height="4"></td>
				</tr>
			</table>

		<%end if
	end if
end if%>

  <table width="100%" border="0" cellpadding="0" cellspacing="0" class=fondoOscuroCblancoAdmin>
    <tr> 
      <td><img src="../global/img/spacer.gif" width="1" height="5"></td>
    </tr>
  </table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><img src="../global/img/spacer.gif" width="1" height="2"></td>
    </tr>
  </table>
<%
	msg_consulta = ""
	if retotal >1 then
		msg_consulta = "Encontrados <b>"& retotal & "</b> registros"
	elseif retotal >0 then
		msg_consulta = "Encontrado <b>1</b> registro"
	else
		msg_consulta = "Ningún registro encontrado"
	end if

	if ""&request.Form("portada")="1" then
		msg_consulta = msg_consulta & " de <b>portada</b>"
	end if

	if ""&nombreSeccionActual <> "" then
		msg_consulta = msg_consulta & " en la sección <b>"& nombreSeccionActual & "</b>"
	end if
	
	if ""&nombreSeccion2Actual <> "" then
		msg_consulta = msg_consulta & " > <b>"& nombreSeccion2Actual & "</b>"
	end if

	if cadena <> "" then
		if inStr(cadena," ")>0 then
			msg_consulta = msg_consulta & " con las palabras <b>"& cadena& "</b>"
		else
			msg_consulta = msg_consulta & " con la palabra <b>"& cadena& "</b>"
		end if
	end if
	
	msg_consulta = msg_consulta & "."
	
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td bgcolor="#F1F3F5"><font color="#849ACE"><%=msg_consulta%></font></td>
  </tr>
</table>


<%if config_descripcion <> "" then%>
	<table width="100%" border="0" cellpadding="1" cellspacing="0" bgcolor="#849ACE">
      <tr>
        <td><table width="100%"  border="0" cellspacing="0" cellpadding="2">
          <tr>
            <td bgcolor="#FFFFFF"><%=config_descripcion%></td>
          </tr>
        </table></td>
      </tr>
    </table>
	<%end if%>

<!-- regInicio:<%=regInicio%> | reTotal: <%=reTotal%> | regLoop: <%=regLoop%> | regporpag: <%=regporpag%><br> -->
  
<%if config_eliminar or config_portada then%>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img src="../global/img/spacer.gif" width="1" height="3"></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 

	<%if config_eliminar and regLoop > 0 then%>
      <td valign="middle"><table  border="0" cellspacing="0" cellpadding="1">
        <tr valign="middle">
          <td width="12">            <table width="100%"  border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><input name="seleccionartodas" id="seleccionartodas" type="checkbox" class="checkReg" onClick="seleccionarTodas(this)" value="checkbox"></td>
              </tr>
            </table></td>
          <td class="txt"><label for="seleccionartodas">Seleccionar todas </label></td>
          <td>&nbsp;</td>
          <td><a id="borrarseleccionadas_btn" href="JavaScript:borrarSeleccion()"><img src="../edicion/img_admin/eliminar.gif" alt=" Borrar seleccionadas " width="13" height="13" border="0" align="absmiddle"></a></td>
          <td class="txt"><label for="borrarseleccionadas_btn">Borrar seleccionadas</label></td>
        </tr>
      </table></td>
	  <%end if%>
	  

	<%if config_portada then%>
      <td align="right">	  <table  border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td class="fondoOscuroAdmin"><table width="100" border="0" cellpadding="0" cellspacing="0" onClick="soloPortada()" style="border: 1px; cursor:hand" title="Mostrar todos los registros que aparecen en portada.">
            <tr>
              <td align="center" valign="middle" bgcolor="#FFFFFF">
				<%if ""&request.Form("portada") = "1" then%>
				<img src="img/bandera_on.gif" width="18" height="18" border="0">
				<%else%>
				<img src="img/bandera.gif" width="18" height="18" border="0">
				<%end if%></td>
              <td align="center" valign="middle" class="fondoAdmin"><nobr> En portada</nobr></td>
            </tr>
          </table></td>
        </tr>
      </table>        </td><%end if%>

    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img src="../global/img/spacer.gif" width="1" height="3"></td>
    </tr>
  </table>
  <%end if%>
  <!-- fin - Seleccionar todas / portada -->

<%
  

	for n=1 to regLoop

		' Rangos de fechas para pintar los colores
		if config_coloresfecha then
			hoy = date()
			r_fecha = ""& re("R_FECHA")
			r_hora = ""& re("R_HORA")
			if inStr(r_hora," ") then
				r_hora = ""
			end if
	
			if ""& r_fecha <> "" and ""& r_fecha <> "0:00:00" and ""& r_hora <> "" then
				fecha_registro = cdate(r_fecha &" "& r_hora)
			elseif ""& r_fecha <> "" then
				fecha_registro = cdate(r_fecha)
			else
				fecha_registro = ""
			end if
			
			if ""& fecha_registro <> "" then
				if hoy<fecha_registro then
					colorfecha = "#9AE492"
				elseif hoy>=fecha_registro and hoy<=fecha_registro then
					colorfecha = "#FFAF95"
				elseif hoy>fecha_registro+1 then
					colorfecha = "#F97575"
				end if
			end if
		end if
%>
			<table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#849ACE" class="fondoOscuroAdmin">
              <tr>
                <td>
                <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="12" rowspan="2" align="center" valign="top" bgcolor="<%=colorfecha%>">
					<table width="12"  border="0" cellpadding="3" cellspacing="0">
                        <tr>
                          <td width="12" height="13" align="center">
						  <%if re("R_ID") <> 1 then%>
						  <input name="marcador<%=n%>" type="checkbox" class="checkReg" value="<%=re("R_ORDEN")%>" idregistro="<%=re("R_ID")%>">
						  <%end if%></td>
                        </tr>
                        <tr>
                          <td align="center">
						<font color="#FFFFFF">
						  <%if idSeccion2Actual > 0 then%>
							  <span title="Sección 2"><%=re("R_ORDEN_SECCION2")%></span>
						  <%elseif idSeccionActual then%>
							  <span title="Sección"><%=re("R_ORDEN_SECCION")%></span>
						  <%else%>
							  <span title="General"><%=re("R_ORDEN")%></span>
						  <%end if%>
						  </font>
						  </td>
                        </tr>
                    </table></td>
                    <td align="left" valign="top" class="fondoAdmin"><table width="100%"  border="0" cellspacing="0" cellpadding="4">
                      <tr>
                        <td align="left" valign="top"><table border="0" align="right" cellpadding="0" cellspacing="0">
                          <tr>
                            <%if re("R_PORTADA") and config_portada then%>
                            <td width="18" height="18"><img src="img/bandera.gif" alt="Este registro aparece en portada" width="18" height="18"></td>
                            <%end if%>
                            <%if config_creador and "no"="seve" then
							if re("R_USUARIO") <> "" then%>
                            <td width="18" height="18"><a href="javascript:alert('Registro creado por: <%=getNombreUsuario(re("R_USUARIO"))%> el día <%=re("R_AUTOFECHA")& " "%><%=re("R_AUTOHORA")%><%if re("R_ULTIMO_USUARIO") <> "" then%>\nEditado por <%=getNombreUsuario(re("R_ULTIMO_USUARIO"))%> el día <%=re("R_ULTIMA_EDICION")%><%end if%>')"><img src="img/usuario.gif" alt=" Usuario " width="18" height="18" border="0" align="absmiddle"></a></td>
                            	<%end if
							end if%>
                            <%if re("R_FOTO") <> "" then%>
                            <td width="18" height="18"><a href="javascript:ampliarfoto('<%=re("R_FOTO")%>')"><img src="img/imagen.gif" width="18" height="18" border="0" align="absmiddle"></a></td>
                            <%end if%>
							<%if config_verenlace then%>
                            <td width="18" height="18"><a href="javascript:verEnlace('<%=re("R_ID")%>')"><img src="img/enlace.gif" alt=" Ver enlace a este registro " width="18" height="18" border="0"></a></td>
							<%end if%>

							<%if config_verenlacefoto and re("R_FOTO") <> "" then%>
                            <td width="18" height="18"><a href="javascript:verEnlaceFoto('<%=re("R_FOTO")%>')"><img src="img/enlacefoto.gif" alt=" Ver link para enlazar foto " width="18" height="18" border="0"></a></td>
							<%end if%>

							<%if config_verenlacearchivo and re("R_ARCHIVO") <> "" then%>
                            <td width="18" height="18"><a href="javascript:verEnlaceArchivo('<%=re("R_ID")%>')"><img src="img/enlacearchivo.gif" alt=" Ver enlace para descargar " width="18" height="18" border="0"></a></td>
							<%end if%>

							<%if config_foro then%>
                            <td width="15" height="18"><a href="JavaScript:respuestas(<%=re("R_ID")%>)"><img src="img/r.gif" width="18" height="18" border="0"></a></td>
							<%end if%>
                            
							<%if idSeccion2Actual > 0 then
								mover = config_mover_seccion2
							elseif idSeccionActual then
								mover = config_mover_seccion
							else
								mover = config_mover
							end if

							if cualid <> "usuarios" then
								if cadena = "" and mover then%>
									<td width="15" height="18"><a href="javascript:moverRegistro('<%=re("R_ID")%>','subir')"><img src="img/flecha_arriba_h.gif" alt=" Subir " width="15" height="18" border="0"></a></td>
									<td width="15" height="18"><a href="javascript:moverRegistro('<%=re("R_ID")%>','bajar')"><img src="img/flecha_abajo_h.gif" alt=" Bajar " width="15" height="18" border="0"></a></td>
								<%else%>
									<td width="15" height="18"><a href="javascript:void 0;"><img src="img/flecha_arriba_h_des.gif" alt=" No es posible ordenar." width="15" height="18" border="0"></a></td>
									<td width="15" height="18"><a href="javascript:void 0;"><img src="img/flecha_abajo_h_des.gif" alt=" No es posible ordenar." width="15" height="18" border="0"></a></td>
								<%end if
							end if%>
                            <%if config_nuevos and re("R_ID") > 1 then%>
							<td width="18" height="18"><a href="javascript:duplicar(<%=re("R_ID")%>)"><img src="img/duplicar.gif" alt="Crear nuevo registro a partir de este. " width="18" height="18" border="0"></a></td>
							<%end if%>
                            <%if config_editar then%>
							<td width="18" height="18"><a href="javascript:editar(<%=re("R_ID")%>,'<%=re("R_SECCION")%>')"><img src="img/lapiz.gif" alt=" Editar " width="18" height="18" border="0" align="absmiddle"></a></td>
							<%end if%>
							<%if config_eliminar and re("R_ID") > 1 then%>
                            <td width="18" height="18"><a href="javascript:eliminar(<%=re("R_ID")%>)"><img src="img/papelera.gif" width="18" height="18" border="0" alt=" Eliminar " align="absmiddle"></a></td>
							<%end if%>
                          </tr>
                        </table>
						<%if config_ampliar then
							Response.Write "<a class=aAdmin href=javascript:ampliar('"& re("R_ID")& "')>"
						end if
						
						if config_idioma_bd = "" then
							Response.Write re("R_TITULO")
						else
							Response.Write re("R_TITULO_" & session("idioma"))
						end if
						
						if config_ampliar then
							Response.Write "</a>"
						end if
						
						if config_fechainifin then
							if ""&re("R_FECHAINI") <> "0:00:00" and ""&re("R_FECHAFIN") <> "0:00:00" then%>
								<br>
								<font color="#849ACE" size="1">Desde <%=re("R_FECHAINI")%> hasta <%=re("R_FECHAFIN")%>.</font>
							<%end if
							
							if ""&re("R_FECHAINI") <> "0:00:00" and ""&re("R_FECHAFIN") = "0:00:00" then%>
								<br>
								<font color="#849ACE" size="1">Comienza <%=re("R_FECHAINI")%>.</font>
							<%end if
							
							if ""&re("R_FECHAINI") = "0:00:00" and ""&re("R_FECHAFIN") <> "0:00:00" then%>
								<br>
								<font color="#849ACE" size="1">Finaliza <%=re("R_FECHAFIN")%>.</font>
							<%end if

						end if%>
						</td>

                      </tr>
                    </table></td>


						<td width="7" rowspan="2" align="left" valign="top" <%
						if config_activo then
							if re("R_ACTIVO") then
								Response.Write "bgcolor='#95E37D' title=' Registro activado '"
							else
								Response.Write "bgcolor='#EA5940' title=' Registro desactivado '"
							end if
						end if

						%>>&nbsp;</td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
                      <tr>
                        <td>
						<%if config_estados then
							if ""& re("R_ESTADO") <> "" then%>
							<img src="../images/<%=re("R_ESTADO")%>.gif" align="right" title="<%=re("R_ESTADO")%>">
							<%end if
						end if%>
						
						<%if re("R_ICONO") <> "" then%>
							<img src="../../datos/<%=session("idioma")%>/<%=cualid%>/iconos/<%=re("R_ICONO")%>" align="right">
						<%end if%>
						
						<%if config_archivo then
							if re("R_ARCHIVO") <> "" then%>
								<%pintaIconoExtension(re("R_TIPOARCHIVO"))%> <a href="../../descargas/?idi=<%=session("idioma")%>&cualid=<%=cualid%>&id=<%=re("R_ID")%>"><%=re("R_ARCHIVO")%></a>
								<br>
							<%end if
						end if%>

						
						<%if config_fecha and re("R_FECHA") <> 0 then%>
							<span class="Estilo5"><%=re("R_FECHA")%></span>
						<%end if%>

						<%if config_cuerpo <> ""then
							if config_idioma_bd = "" then
								Response.Write unpoco(re("R_"& config_cuerpo),330)
							else
								Response.Write unpoco(re("R_"& config_cuerpo & "_" & session("idioma")),330)
							end if
						end if%></td>
                      </tr>
                    </table>
					</td>
                  </tr>
				  
                </table></td>
              </tr>
      </table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img src="../global/img/spacer.gif" width="1" height="4"></td>
    </tr>
  </table>
  <%re.movenext
		next%>

  <%if not regLoop >0 then%>
	  <br>
<div align="center"><b>No se han encontrado resultados</b></div>
  <%end if%>		

  <!-- PAGINADO -->
  <%if regLoop >0 then%>
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
      <td width="100%" align="center"><table width="200" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#EBEFF5">
        <tr>
          <td align="left"><%if regInicio > 1 then %>
              <input name="" type="button" class="botonAdmin" onClick="goPag('-')" value="<<">
              <%else%>
              <input name="" type="button" disabled class="botonApagadoAdmin" value="<<">
              <%end if%>
          </td>
          <td align="center"><%
	dim c, cmax
	dim numlinkpag
	numlinkpag = 5
	for n=1 to retotal
		if n mod regporpag = 1 then
			cmax = cmax + 1
		end if
	next
	if cmax > 1 then
		for n=1 to retotal
			if n mod regporpag = 1 then
				c = c + 1
				if (n < regInicio+(regporpag*numlinkpag) and n > regInicio-(regporpag*numlinkpag)) then
					if regInicio = n then%>
      [<%=c%>]
      <%else%>
      <a href="#" onClick="goPag('<%=n%>')"><%=c%></a>
      <%end if
				end if
			end if
		next
	end if%>
          </td>
          <td align="right"><%if regInicio+regPorPag <= retotal then%>
              <input name="" type="button" class="botonAdmin" onClick="goPag('+')" value=">>">
              <%else%>
              <input name="" type="button" disabled="true" class="botonApagadoAdmin" value=">>">
              <%end if%></td>
        </tr>
      </table></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <br>
<%end if%>

<span title="Consulta acutal:
<%=alt_sql%>">[?]</span>

  <!-- FIN: PAGINADO -->
  <%consultaXClose()
	end if ' unerror%>

	
  <!-- <a href="#" onClick="alert(&quot;<%=pintaSqlJS(sql)%>&quot;)">VER SQL</a> -->
  <%end select%>

  <%if unerror then%>
 <b>Error:</b><br><%=msgerror%>
  <%end if%>
</form>
<script language="javascript" type="text/javascript">
	<!--
	//
	function changeAlfinal(c){
		//try {
			nuevosalfinal = parent.frames[1].f.nuevosalfinal
			if(c.checked){
				nuevosalfinal.value = 1
			} else {
				nuevosalfinal.value = 0
			}
		//}catch(unerror){}
		
	}
	//
	function aplicarColor(nombre,color){
		if(""+ nombre != "") {
			f[nombre].value = color
		}
	}
	//
	function noAnadirFecha(){
		f.fecha_dia.disabled = true
		f.fecha_mes.disabled = true
		f.fecha_ano.disabled = true
	}
	function siAnadirFecha(){
		f.fecha_dia.disabled = false
		f.fecha_mes.disabled = false
		f.fecha_ano.disabled = false
	}
	
	function noAnadirFechaini(){
		f.fechaini_dia.disabled = true
		f.fechaini_mes.disabled = true
		f.fechaini_ano.disabled = true
	}
	function siAnadirFechaini(){
		f.fechaini_dia.disabled = false
		f.fechaini_mes.disabled = false
		f.fechaini_ano.disabled = false
	}
	
	function noAnadirFechafin(){
		f.fechafin_dia.disabled = true
		f.fechafin_mes.disabled = true
		f.fechafin_ano.disabled = true
	}
	function siAnadirFechafin(){
		f.fechafin_dia.disabled = false
		f.fechafin_mes.disabled = false
		f.fechafin_ano.disabled = false
	}

	//
	function soloPortada() {
		try {
			f.ac.value = ""
			f.action = "#"
			f.target = ""
			<%if ""&request.Form("portada") = "1" then%>
			f.portada.value = 0
			<%else%>
			f.portada.value = 1
			<%end if%>
			f.submit()
		}catch(unerror){}
	}
	//
	function editarOrdenIdioma(id) {
		location.href='main.asp?ac=editarordenidioma&id='+id
	}
	//
	function editarSeccion(id) {
		location.href='main.asp?ac=editarseccion&id='+id
	}
	//
	function editarSeccion2(id,seccion) {
		location.href='main.asp?ac=editarseccion2&seccion='+ seccion +'&id='+id
	}
	//
	function eliminarOrdenIdioma(id) {
		if(confirm("¿Desea eliminar este registro?")){
			location.href='main.asp?ac=eliminarordenidioma&id='+id
		}
	}
	//
	function eliminarSeccion(id) {
		if(confirm("¿Desea eliminar esta sección?")){
			location.href='main.asp?ac=eliminarseccion&id='+id
		}
	}
	//
	function eliminarSeccion2(id,seccion) {
		if(confirm("¿Desea eliminar esta sub-sección?")){
			location.href='main.asp?ac=eliminarseccion2&seccion='+ seccion +'&id='+id
		}
	}

	// Ver enlace
	function verEnlace(id){
		ventana("main.asp?ac=verenlace&id="+id,"VerEnlace",500,150,0)
	}
	
	// Ver enlace archivo
	function verEnlaceArchivo(id){
		ventana("main.asp?ac=verenlacearchivo&id="+id,"VerEnlaceArchivo",500,150,0)
	}
	
	// Ver enlace foto
	function verEnlaceFoto(foto){
		ventana("main.asp?ac=verenlacefoto&foto="+foto,"VerEnlaceFoto",500,150,0)
	}

	// Mover registro
	function moverRegistro(id,dire){
		try{
			f.ac.value = "moverregistro"
			f.id.value = id
			f.comun.value = dire
			f.target = "der"
			f.submit()			
		}catch(unerror){
			//
		}
	}
	// Mover orden idioma
	function moverOrdenIdioma(id,dire){
		try{
			var f = document.getElementById("f");
			f.ac.value = "moverordenidioma"
			f.id.value = id
			f.comun.value = dire
			f.submit()			
		}catch(unerror){}
	}

	// Mover sección
	function moverSeccion(id,dire){
		try{
			var f = document.getElementById("f");
			f.ac.value = "moverseccion"
			f.id.value = id
			f.comun.value = dire
			f.submit()			
		}catch(unerror){
			alert(unerror.description)
		}
	}
	function moverSeccion2(id,dire){
		try{
			var f = document.getElementById("f");
			f.ac.value = "moverseccion2"
			f.id.value = id
			f.comun.value = dire
			f.submit()			
		}catch(unerror){
			alert(unerror.description)
		}
	}
	
	// Intercambia el orden de un registro con el de otro
	function iOrden(r) {
		ventana("main.asp?ac=iorden&r="+r,'IntercambiarOrden',300,150,0)
	}
	//
	function iorden_ya(pos1,pos2){
		try{
			if (pos1 == pos2){
				alert("Escoja posiciones distintas.")
				f.pos2.focus()
			}else{
				f.disabled = true
				location.href="main.asp?ac=iorden&pos1="+pos1+"&pos2="+pos2
			}
		}catch(unerror){
		}
	}
	// Duplicar un registro
	function duplicar(id) {
		try{
			f.ac.value = "duplicar"
			f.id.value = id
			f.target = "der"
			f.submit()
		}catch(unerror){}
	}
	// Eliminar un registro
	function eliminar(id) {
		try{
			if(confirm("¿Está seguro que desea eliminar este registro?")) {
				f.ac.value = "eliminar"
				f.id.value = id
				f.target = "der"
				f.submit()			
			}
		}catch(unerror){
			//
		}
	}
	//
	// borrarSeleccion()
	function borrarSeleccion() {
		var alguna 
		var registros = ""
		var cadena = ""
		var check
		for (var n=1;n<=<%=retotal%>;n++){
			check = f["marcador" + n]
			if (check)
				if (check.checked == true) cadena = cadena + f["marcador" + n].idregistro + ",";
		}
		cadena = cadena + "fin"

		if (cadena != "fin" && cadena != "") {
			if(confirm("¿Seguro que desea borrar los registros seleccionados.")){
				f.ac.value = "borrarseleccion"
				f.target = "der"
				f.comun.value = cadena
				f.submit()
			}
		} else {
			alert("Seleccione al menos un registro para borrar.")
		}

	}
	//
	// Administrar Orden de idioma
	function adminOrdenIdioma() {
		try{
			f.ac.value = "adminordenidioma"
			f.target = "der"
			f.submit()
			f.ac.value = ""
		}catch(unerror){
			//
		}
	}

	//
	// Administrar secciones
	function adminSecciones() {
		try{
			f.ac.value = "adminsecciones"
			f.target = "der"
			f.submit()
			f.ac.value = ""
		}catch(unerror){}
	}
	//
	// Administrar secciones2 (Sub-Secciones)
	function adminSecciones2() {
		try{
			f.ac.value = "adminsecciones2"
			f.target = "der"
			f.submit()
			f.ac.value = ""
		}catch(unerror){}
	}
	//
	// Editar
	function editar(id,seccion) {
		try{
			f.ac.value = "editar"
			f.action = "main.asp"
			f.target = "der"
			f.id.value = id
			f.comun.value = seccion // sección del registro actual
			f.submit()
			f.ac.value = ""
		}catch(unerror){}
	}
	//
	// Ver todas las respuesta a este tema
	function respuestas(id){
		f.action = "foro.asp"
		f.ac.value = "respuestas"
		f.target = "der"
		f.id.value = id
		f.submit()
		f.action = "main.asp"
	}
	//
	// Ampliar un registro
	function ampliar(id) {
		try{
			f.ac.value = "ampliar"
			f.target = "der"
			f.id.value = id
			f.submit()
			f.ac.value = ""
			f.target = ""
		}catch(unerror){
			//
		}
	}
	//
	// seleccionarTodas
	function seleccionarTodas(c) {
		try {
			for (var n=1;n<=<%=retotal%>;n++){
				f["marcador" + n].checked = c.checked
			}
		} catch(unerror){}
	}
	//
	//
	function ampliarfoto(nombre) {
		ventana("archivos.asp?ac=ampliarfoto&archivo="+nombre,'AmpliarFoto',100,100,0)
	}
	function ampliaricono(nombre) {
		ventana("archivos.asp?ac=ampliaricono&archivo="+nombre,'AmpliarIcono',100,100,0)
	}
	//
	function borrarCadena() {
		try {
			f.cadena.value = ""
			f.reginicio.value = ""
			f.submit()
		} catch(unerror) {
			//
		}
	}
	//
	// Cambio de sección
	function changeSeccion( c ) {
		var f = document.f;
		f.ac.value = "";
		f.target="";
		if (c.value != "") {
			if(f.seccion2) f.seccion2.value = "";
			f.reginicio.value = "";
			f.seccion.value = c.value;
			f.submit();
		}
	}
	//
	// Cambio de estado
	function changeEstado( c ) {
		document.f.submit();
	}
	//
	// Cambio de sección2
	function changeSeccion2(c) {
		f.ac.value = ""
		f.target=""
		try {
			f.reginicio.value = ""
			f.seccion2.value = c.value
			f.submit()
		} catch(unerror){
			alert(unerror)
		}
	}
	//
	function nuevo() {
		try {
			var unerror = false

			<%if config_idioma_bd <> "" then%>
				alert("Los datos que usted está manejando desde este idioma están fisicamente en la base de datos del idioma principal (Español).\nPor favor, para insertar nuevos registros debe ir al idioma principal.");
				unerror = true
			<%end if%>

			<%if ""&launica = "" then%>
			if (""+f.seccion.value == "" && !unerror){
				alert("Por favor, Escoja una sección.")
				f.seccion.focus()
				unerror = true
			}
			<%end if%>

			if(!unerror){
				<%if ""&launica <> "" then%>
					f.seccion.value = <%=launica%>
				<%end if%>
				f.target="der"
				f.ac.value = "nuevo"
				f.comun.value = ""
				f.cadena.value = ""
				f.submit()
				f.target=""
				f.ac.value = ""
				<%if ""&launica <> "" and idSeccionActual = 0 then%>
					f.seccion.value = ""
				<%end if%>
			}
		} catch (unerror){}
	}
	//
	function changeRegPorPag(v) {
		try{
			f.ac.value = "main.asp"
			f.reginicio.value = ""
			f.target=""
			f.submit()
		} catch(unerror){
			//
		}
	}
	//
	// Ir a una pagina
	function goPag(s) {
		try{
			f.ac.value = "main.asp"
			f.target=""
			if(f.cadena.value != f.cadenabak.value) {
				f.reginicio.value = ""
			}
			if (s == "+") {
				f.reginicio.value = Number(f.reginicio.value) + <%=regporpag%>
			} else if(s == "-") {
				f.reginicio.value = Number(f.reginicio.value) - <%=regporpag%>
			} else {
				f.reginicio.value = s
			}
			f.submit()
		} catch(unerror){
			//
		}
	}
	//
	function envio() {
		er_telefono = /(^([0-9]{9,13})|^)$/;
		er_numero = /(^([0-9])|^)$/;
		var er_email = /^(.+\@.+\..+)$/;

		f.target=""
		try{
			if(f.cadena.value != f.cadenabak.value) {
				f.reginicio.value = ""
			}
		} catch(unerror){
			//alert(unerror.description)
		}

		// VALIDACIÓN  DE CAMPOS **********************************************
		
		<%if cualid = "usuarios" and numero(seccion) >0 then%>
		var cambiar_clave
		cambiar_clave = f.cambiar_clave;
		if (cambiar_clave){
			if (cambiar_clave.checked){
				if (""+f.clave.value == ""){
					alert("Escriba la nueva clave.\nSi no desea cambiar la clave desactive la casilla 'Cambiar clave'.")
					f.clave.focus()
					return false
				} else {
					if (f.clave.value != f.clave_r.value) {
						alert("Las claves escritas no coinciden.")
						f.clave.value = ""
						f.clave_r.value = ""
						f.clave.focus()
						return false
					}
				}
			}
		}
		<%end if%>

		<%if ac = "editar" or ac = "nuevo" then%>
		// Fijos (de arriba):
		if (f.seccion.value == ""){
			alert("Por favor, Escoja una sección.")
			f.seccion.focus()
			return false
		}
		<%if config_activo_seccion2 then%>
		if (f.seccion2){
			if (f.seccion2.value == 1){
				alert("Por favor, Escoja una subsección.")
				f.seccion2.focus()
				return false
			}
		}
		<%end if%>
		if (f.titulo.value == ""){
			alert("Por favor, rellene el campo \"<%=config_nom_titulo%>\".")
			f.titulo.focus()
			return false
		}
		
		// Dinámicos:
		<%
		if cualid = "usuarios" and seccion >0 then
			set miGrupo = setGrupo(seccion)
			if typeOK(miGrupo) then
				set nodosCampo = miGrupo.selectNodes("//grupos/grupo[@id="& seccion &"]//dato")
			else
				unerror = true : msgerror = "No se ha encontrado el grupo en el XML."
			end if
		else
			set nodosCampo = nodoCualid.childNodes
		end if
		
		for each a in nodosCampo
			if a.nodeName = "dato" then
				nombrecorto = a.getAttribute("nombrecorto")
				c_titulo = a.getAttribute("titulo")
				c_nombre = a.getAttribute("nombre")
				c_tipo = a.getAttribute("tipo")
			
				' EMAIL '
				if ""&a.getAttribute("validar") = "email" then
					if ""&a.getAttribute("requerido") = "1" then
						%>
						if(f.<%=nombrecorto%>.value == "") {
							alert('Por favor, rellene el campo \"<%=c_titulo%>\".');
							f.<%=nombrecorto%>.focus();
							return false;
						}
						if(!er_email.test(f.<%=nombrecorto%>.value)) {
							alert('Por favor, introduzca una dirección E-mail válida en el campo \"<%=c_titulo%>\".');
							f.<%=nombrecorto%>.focus();
							return false
						}
						<%
					else
						%>
						if(f.<%=nombrecorto%>.value != "") {
							if(!er_email.test(f.<%=nombrecorto%>.value)) {
								alert('Introduzca una dirección E-mail válida en el campo \"<%=c_titulo%>\" o déjelo vacio.');
								f.<%=nombrecorto%>.focus();
								return false
							}
						}
						<%
					end if
				
				' NUMÉRICO '
				elseif ""&a.getAttribute("validar") = "numero" then
					if ""&a.getAttribute("requerido") = "1" then
						%>
						if(f.<%=nombrecorto%>.value == "") {
							alert('Por favor, rellene el campo \"<%=c_titulo%>\".');
							f.<%=nombrecorto%>.focus();
							return false;
						}
						if(isNaN(f.<%=nombrecorto%>.value)) {
							alert('Por favor, rellene el campo \"<%=c_titulo%>\" con un valor numérico.');
							f.<%=nombrecorto%>.focus();
							return false;
						}
						<%
					else
						%>
						if(f.<%=nombrecorto%>.value != "") {
							if(isNaN(f.<%=nombrecorto%>.value)) {
								alert('Rellene el campo \"<%=c_titulo%>\" con un valor numérico o déjelo vacio.');
								f.<%=nombrecorto%>.focus();
								return false;
							}
						}
						<%
					end if
				
				' REQUERIDO (cualquier tipo de campo)'
				elseif ""&a.getAttribute("requerido") = "1" then
					select case ""&a.getAttribute("tipo")
					case "texto"
						%>
						if(f.<%=nombrecorto%>.value == "") {
							alert('Por favor, rellene el campo \"<%=c_titulo%>\".');
							f.<%=nombrecorto%>.focus();
							return false;
						}
						<%
					case "opcion"
						%>
						ok = false
						for(var n=1;n<=f.<%=nombrecorto%>.length;n++){
							if (f.<%=nombrecorto%>[n-1].checked)
								ok = true
						}
						if(ok == false) {
							alert('Por favor, seleccione una opción para \"<%=c_titulo%>\".');
							//f.<%=nombrecorto%>.focus(); (no admite foco ...)
							return false;
						}
						<%
					case "memo"
						%>
						if(f.<%=nombrecorto%>.value == "") {
							alert('Por favor, rellene el campo \"<%=c_titulo%>\".');
							<%if ""&a.getAttribute("editorhtml") <> "1" then%>
								f.<%=nombrecorto%>.focus();
							<%end if%>
							return false;
						}
						<%
					end select
				end if
			end if
		next

	end if%>
		// ******************************************************************************


	}
	// Ver todo
	function verTodo(){
		try{
			f.target=""
			f.reginicio.value = ""
			f.ac.value = ""
			f.cadena.value = ""
			f.seccion.value = ""
			f.portada.value = ""
			f.estado.value = ""
			f.submit()
		} catch(unerror){
			//
		}
	}
	
	
//-->
</script>
	
<%end if

	sub numVer(n)
		if ""&n <> "" then%>
		<table width="7" border="0" cellspacing="0" cellpadding="0">
		<%for v=1 to len(n)%>
			<tr>
			<td width="7" height="10" align="center"><img src="img/numeros/n<%=mid(n,v,1)%>.gif" width="7" height="10"></td>
			</tr>
		<%next%>
		</table>
		<%end if
	end sub
	
	sub pintaIconoExtension (ext)
		select case ext
			case "jpg","png","gif","bmp","tif"
				ico = "img"
			case "exe"
				ico = "exe"
			case "doc"
				ico = "doc"
			case "txt"
				ico = "txt"
			case "xls"
				ico = "xls"
			case "zip"
				ico = "zip"
			case "pdf"
				ico = "pdf"
			case "mp3","wav"
				ico = "mp3"
			case else
				ico = "iconootro"
		end select
			%><img src="../../img/<%=ico%>.gif" align="absmiddle" alt=" Archivo tipo: <%=UCASE(ext)%> "><%
	end sub


id = ""&request.QueryString("id")
direct = ""&request.QueryString("direct")
vengoDe = inStr(request.ServerVariables("HTTP_REFERER"),"global/aSkipper.asp")
seccion = ""&request.QueryString("seccion")

	if direct <> "" and vengoDe >0 then%>
		<script>
			switch ("<%=direct%>") {
			  case "ampliar":
				ampliar(<%=id%>)
				break;
			  case "editar":
				editar('<%=id%>','<%=seccion%>')
				break;
			  case "eliminar":
				eliminar('<%=id%>')
				break;
			}
		</script>
	<%end if%>
</body>
</html>