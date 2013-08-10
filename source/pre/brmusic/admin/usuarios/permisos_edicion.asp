<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/usuarios/rutinasParaAdmin.asp" -->
<html>
<head>
<title>aSkipper</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global/estilos.css" rel="stylesheet" type="text/css">
</head>
<body vlink="#003366" class="bodyAdmin">
<%if unerror then
	Response.Write "<b>Error</b><br>" & msgerror
else%>
<span class="tituloazonaAdmin">Permisos para secciones</span><br>
<br>
Escoja las secciones que desea administrar.<br>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="1" bgcolor="#CCCCCC"><img src="../../spacer.gif" width="1" height="1"></td>
      </tr>
    </table>	
<%
	dim permisosActual
	permisosActual = replace(replace(""&request.QueryString("P"),"/","-"),"|","X")

	Function elemento(cadena,numero)
		coincidencia = 0
		' paso las "/"
		for n=0 to numero-1
			coincidencia = Instr(coincidencia+1,cadena,"/")
		next
		resto = Instr(coincidencia+1,cadena,"/")
		'elemento=cadena&resto-coincidencia-1
		if resto>0 then
			elemento = Mid(cadena,coincidencia+1,resto-coincidencia-1)
		else
			elemento = "noencontrado"
		end if
	end Function
			
	function pintaruta(valor,elxml)
		dim n,rutaxml,devuelto
		set rutaxml = elxml
		for n=1 to valor
			Set rutaxml = rutaxml.parentnode
			devuelto = rutaxml.nodename&"/"&devuelto
		next
		pintaruta = devuelto
	end function

	function activarSeleccionado(pLugar,pSeleccionActual)
		dim l
		l = "X" & replace(pLugar,"/","-") & "X"
		if inStr(pSeleccionActual,l) or inStr(permisosActual,l) then
			activarSeleccionado = "checked"
		end if
	end function

	Sub ver(Nodos,seccion)
		Dim oNodo
		pos = 1
		
		nhijos = 0
		For Each oNodo In Nodos
			nhijos = nhijos + 1
			if oNodo.nodeType = 1 then
				xmlfile = oNodo.nodename&".xml"
				Set anterior=oNodo
				rutanterior = ""
				
				for n=0 to separador
					rutanterior=anterior.nodename&"/"&rutanterior
					Set anterior=anterior.parentnode
				next
				
				lugar = pintaruta(separador,oNodo) & oNodo.nodename
				voyapintar=5
				
				for separa=0 to separador
					if strComp(elemento(request("desplegadoActual"),separa),elemento(rutanterior,separa))=0 then
						voyapintar=voyapintar+1
					end if
				next
				
				if separador=0 or voyapintar>=separa-1  then%>
				<table width="100%"  border="0" cellpadding="0" cellspacing="0">
					<tr>
					<td width="15">&nbsp;</td>
					<td>
					<table width="100%" border="0" cellpadding="0" cellspacing="1">
						<tr>
						<%
						editable = Cbool(0 + ("0" & oNodo.getattribute("editable")))
						if strComp(oNodo.parentnode.nodename,"secciones") = 0 then%>
							<td>
							<table  border="0" cellspacing="0" cellpadding="1">
								<tr>
								<%if editable then%>
									<td><input name="<%=replace(oNodo.parentNode.nodeName,"/","_") & "_" & nhijos%>" id="<%=replace(lugar,"/","_") & "_" & nhijos%>" type="checkbox" lugar="<%=lugar%>" class="checkPermisos" onClick="seleccion(this,'<%=lugar%>','<%=oNodo.nodeName%>')" value="checkbox" <%=activarSeleccionado(lugar,request.Form("seleccionActual"))%>></td>
								<%end if
								if oNodo.childnodes.length>1 and "esta" = "anulado" then
									if voyapintar <= separador then%>
										<td valign="middle"><a href="JavaScript:desplegar('<%=lugar%>/')"><img src="../images/mas.gif" width="9" height="9" border="0" align="absmiddle"></a></td>
									<%else%>
										<td><a href="JavaScript:plegar('<%=Left(lugar,InstrRev(lugar,"/"))&"/"%>')"><img src="../images/menos.gif" border="0" align="absmiddle"></a></td>
									<%end if
								end if%></tr>
							  </table></td>
							<td width="100%" bgcolor="#f9f9f9"><%if oNodo.childnodes.length>1 and voyapintar > separador then Response.Write "<b>"& oNodo.getattribute("titulo") &"</b>" else Response.Write oNodo.getattribute("titulo") end if%></td>
						<%elseif separador=1 then%>
						<td>
							<table  border="0" cellspacing="0" cellpadding="1">
								<tr>
								<td valign="middle"><img src="../images/espacio.gif" width="9" height="9"></td>

								<%if oNodo.childnodes.length>1 then
									if voyapintar<=separador then
										estaDesplegado = 1
									else
										estaDesplegado = 2
									end if
								else
									estaDesplegado = 0
								end if%>

								<%if editable then%>
									<td><input name="<%=replace(oNodo.parentNode.nodeName,"/","_") & "_" & nhijos%>" id="<%=replace(lugar,"/","_") & "_" & nhijos%>" type="checkbox" lugar="<%=lugar%>" class="checkPermisos" onClick="seleccion(this,'<%=lugar%>','<%=oNodo.nodeName%>')" value="checkbox" <%=activarSeleccionado(lugar,request.Form("seleccionActual"))%>></td>
								<%end if%>
								<%if estaDesplegado > 0 and "está" = "anulado" then
									if estaDesplegado = 1 then%>
										<td valign="middle"><a href="JavaScript:desplegar('<%=lugar%>/')"><img src="../images/mas.gif" width="9" height="9" border="0" align="absmiddle"></a></td>
									<%elseif estaDesplegado = 2 then%>
										<td><a href="JavaScript:plegar('<%=Left(lugar,InstrRev(lugar,"/"))&"/"%>')"><img src="../images/menos.gif" border="0" align="absmiddle"></a></td>
									<%end if
								end if%>
								<%if editable then%>
									<!--<td><input name="checkbox" type="checkbox" class="checkPermisos" onClick="seleccion(this,'<%=lugar%>')" value="checkbox" <%=activarSeleccionado(lugar,request.Form("seleccionActual"))%>></td>-->
								<%end if%></tr>
						  </table></td>
							<td width="100%"><%if oNodo.childnodes.length>1 and voyapintar > separador then Response.Write "<b>"& oNodo.getattribute("titulo") &"</b>" else Response.Write oNodo.getattribute("titulo") end if%></td>
						<%elseif separador=2 then%>
							<td>
							<table  border="0" cellspacing="0" cellpadding="1">
								<tr>
								<td valign="middle"><img src="../images/espacio.gif" width="9" height="9"></td>
								<td valign="middle"><img src="../images/espacio.gif" width="9" height="9"></td>
								<%if editable then%>
									<td><input name="<%=replace(oNodo.parentNode.nodeName,"/","_") & "_" & nhijos%>" id="<%=replace(lugar,"/","_") & "_" & nhijos%>" type="checkbox" lugar="<%=lugar%>" class="checkPermisos" onClick="seleccion(this,'<%=lugar%>','<%=oNodo.nodeName%>')" value="checkbox" <%=activarSeleccionado(lugar,request.Form("seleccionActual"))%>></td>
								<%end if%>
								<%if oNodo.childnodes.length>1 then
									if voyapintar <= separador then%>
										<td valign="middle"><a href="JavaScript:desplegar('<%=lugar%>/')"><img src="../images/mas.gif" width="9" height="9" border="0" align="absmiddle"></a></td>
									<%else%>
										<td><a href="JavaScript:plegar('<%=Left(lugar,InstrRev(lugar,"/"))&"/"%>')"><img src="../images/menos.gif" border="0" align="absmiddle"></a></td>
									<%end if
								end if%></tr>
							</table></td>
								<td width="100%"><%if oNodo.childnodes.length>1 and voyapintar > separador then Response.Write "<b>"& oNodo.getattribute("titulo") &"</b>" else Response.Write oNodo.getattribute("titulo") end if%></td>
							<%elseif separador=3 then%>
								<td>
								<table border="0" cellspacing="0" cellpadding="1">
									<tr valign="middle">
									<td><img src="../images/espacio.gif" width="9" height="9"></td>
									<td><img src="../images/espacio.gif" width="9" height="9"></td>
									<td><img src="../images/espacio.gif" width="9" height="9"></td>
									<td><input name="<%=replace(oNodo.parentNode.nodeName,"/","_") & "_" & nhijos%>" id="<%=replace(lugar,"/","_") & "_" & nhijos%>" type="checkbox" lugar="<%=lugar%>" class="checkPermisos" onClick="seleccion(this,'<%=lugar%>','<%=oNodo.nodeName%>')" value="checkbox" <%=activarSeleccionado(lugar,request.Form("seleccionActual"))%>></td>
									</tr>
								</table></td>
								<td width="100%"><%=oNodo.getattribute("titulo")%></td>
							<%end if%>
					  </tr>
					  </table></td>
						<td width="20">&nbsp;</td>
				  </tr>
</table>
					<%else 
				end if
			end if
			
			If oNodo.hasChildNodes Then
				separador=separador+1
				ver oNodo.childNodes,oNodo.nodename
				separador=separador-1
			End If
		Next
	End Sub

%>
<form name="f" action="#" method="post">
	<input name="desplegadoActual" type="hidden" value="<%=request.Form("desplegadoActual")%>">
	<input name="seleccionActual" type="hidden" id="seleccionActual" value="<%=request.Form("seleccionActual")%>">
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
		<td>
		<%
		Dim valorant, separador, xmlObj
		separador = 0
		Set xmlObj = CreateObject("MSXML.DOMDocument")
		if xmlObj.Load(Server.MapPath("../../"&session("idioma")&"/secciones.xml")) then
			ver xmlObj.selectsinglenode("/pagina/secciones").childnodes,"principal"
		Else
			response.Write("Ha ocurrido un error.")
		End If
		set FSO=Nothing
		%></td>
		</tr>
  </table>
    <br>
    <hr width="100%" size="1" noshade>
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="right"><input name="" type="button" class="botonAdmin" onClick="aceptar()" value="Aceptar"></td>
      </tr>
    </table>
</form>
<script language="javascript" type="text/javascript">
<!--
	var str_seleccionados
	if(""+f.seleccionActual.value == ""){
		if (parent.opener.f.secciones_esp.value != ""){
			f.seleccionActual.value = str_seleccionados = parent.opener.f.secciones_esp.value.toLowerCase().replace(/\|/g,"X").replace(/\//g,"-")
		} else {
			str_seleccionados = "X"
		}
	}else{
		str_seleccionados = f.seleccionActual.value
	}
	
	//
	function cadenaActivos(c){
		lugar = c.lugar.replace(/\//g,"-");
		if (c.checked){
			if (str_seleccionados==""){
				str_seleccionados = "X"+ lugar +"X";
				
			}else{
				str_seleccionados = str_seleccionados.replace(eval("/X"+lugar+"X/g"),"X");
				str_seleccionados += lugar +"X";
			}
		}else{
			str_seleccionados = str_seleccionados.replace(eval("/X"+lugar+"X/g"),"X");
		}
		str_seleccionados = str_seleccionados.replace(/\X\X/g,"X");
		str_seleccionados = str_seleccionados.replace(/^\X$/g,"");
		f.seleccionActual.value = str_seleccionados;

	}
	function seleccion(c,yo){
		for(var h=0; h<f.length; h++){
			if(f.item(h).id.indexOf(yo+"_") >=0){
				var campo = document.getElementById(f.item(h).id)
				campo.checked = c.checked
				cadenaActivos(campo)
			}
		}
		cadenaActivos(c)
	}
	
	//
	function desplegar(lugar){
		f.desplegadoActual.value = lugar;
		f.submit();
	}
	
	//
	function plegar(lugar){
		f.desplegadoActual.value = lugar;
		f.submit();
	}
	
	// 
	function aceptar(){
		var fp = parent.opener.f;
		if (""+fp == "[object]"){
			str_seleccionados = str_seleccionados.replace(/-/g,"/");
			str_seleccionados = str_seleccionados.replace(/X/g,"|");
			fp.secciones_esp.value = str_seleccionados;
		} else {
			alert("No se ha encontrado la ventana principal.\nRepita el proceso.");
		}
		window.close();
	}

	//-->
</script>
<%end if ' unerror%>

</body>
</html>