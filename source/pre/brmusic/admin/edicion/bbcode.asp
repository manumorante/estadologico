<script language="Javascript1.2">
<!--
// -------------------------------------------------------------------------------  htmlarea
_editor_url = "";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
	document.write('<scr' + 'ipt src="' +_editor_url+ 'Editor.js"');
	document.write(' language="Javascript1.2"></scr' + 'ipt>');
} else {
	document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>');
}

misBotones = 0


//-------------------------------------------------------------------------------------------------------- -->
</script>

<script src="../rutinas.js" type="text/javascript" language="JavaScript"></script>
<script language="JavaScript" type="text/javascript">
<!--
// variable global = focoActual
var focoActual = ""
var unError = false

function ventana(theURL,winName,ancho,alto,features) { 
	var winl = (screen.width - ancho) / 2;
	var wint = (screen.height - alto) / 2;
	var paramet=features+',top='+wint+',left='+winl+',width='+ancho+',height='+alto;
	var splashWin=window.open(theURL,winName,paramet);
    splashWin.focus();
}


// bbCode control by
// subBlue design
// www.subBlue.com

// Startup variables
var imageTag = false;
var theSelection = false;

// Check for Browser & Platform for PC & IE specific bits
// More details from: http://www.mozilla.org/docs/web-developer/sniffer/browser_type.html
var clientPC = navigator.userAgent.toLowerCase(); // Get client info
var clientVer = parseInt(navigator.appVersion); // Get browser version

var is_ie = ((clientPC.indexOf("msie") != -1) && (clientPC.indexOf("opera") == -1));
var is_nav  = ((clientPC.indexOf('mozilla')!=-1) && (clientPC.indexOf('spoofer')==-1)
                && (clientPC.indexOf('compatible') == -1) && (clientPC.indexOf('opera')==-1)
                && (clientPC.indexOf('webtv')==-1) && (clientPC.indexOf('hotjava')==-1));

var is_win   = ((clientPC.indexOf("win")!=-1) || (clientPC.indexOf("16bit") != -1));
var is_mac    = (clientPC.indexOf("mac")!=-1);


// Definición de etiquetas
bbcode = new Array();
//bbtags = new Array('[b]','[/b]','[i]','[/i]','[u]','[/u]','[quote]','[/quote]','[code]','[/code]','[list]','[/list]','[list=]','[/list]','[img]','[/img]','[url]','[/url]');
bbtags = new Array('<b>','</b>','<i>','</i>','<u>','</u>','','<hr>','<ul><li>','</li></ul>','[url]','[/url]','','');
imageTag = false;

// Replacement for arrayname.length property
function getarraysize(thearray) {
	for (i = 0; i < thearray.length; i++) {
		if ((thearray[i] == "undefined") || (thearray[i] == "") || (thearray[i] == null))
			return i;
		}
	return thearray.length;
}

// Replacement for arrayname.push(value) not implemented in IE until version 5.5
// Appends element to the array
function arraypush(thearray,value) {
	thearray[ getarraysize(thearray) ] = value;
}

// Replacement for arrayname.pop() not implemented in IE until version 5.5
// Removes and returns the last element of an array
function arraypop(thearray) {
	thearraysize = getarraysize(thearray);
	retval = thearray[thearraysize - 1];
	delete thearray[thearraysize - 1];
	return retval;
}

function bbstyle(bbnumber,nombreCampo,nombreForm) {
	donotinsert = false;
	theSelection = false;
	bblast = 0;

	if (bbnumber == -1) { // Close all open tags & default button names
		while (bbcode[0]) {
			butnumber = arraypop(bbcode) - 1;
			document[nombreForm][nombreCampo].value += bbtags[butnumber + 1];
			buttext = eval('document[nombreForm].addbbcode' + butnumber + '.value');
			eval('document[nombreForm].addbbcode' + butnumber + '.value ="' + buttext.substr(0,(buttext.length - 1)) + '"');
		}
		imageTag = false; // All tags are closed including image tags :D
		document[nombreForm][nombreCampo].focus();
		return;
	}

	if ((clientVer >= 4) && is_ie && is_win)
		theSelection = document.selection.createRange().text; // Get text selection
	if (theSelection) {
		// Add tags around selection
		// COMPROBACIÓN PARA QUE NO SE SELECCIONE EL CAMPO ENTERO, EVITANDO DESTRUIRLO
		valorDelCampo = document[nombreForm][nombreCampo].value
		elIndex = valorDelCampo.indexOf(theSelection)
		cc = valorDelCampo.indexOf(theSelection)
		if (focoActual == nombreCampo) {
			// le reemplazo al value del campo la Seleccion por la seleccion con los cambios.
			if (cc>=0 && theSelection!='' && theSelection!=' ') {
				if (bbnumber == 12) {
					document.selection.createRange().text = theSelection.toUpperCase()
				}else if (bbnumber == 14){
					document.selection.createRange().text = theSelection.toLowerCase()
				}else{
					document.selection.createRange().text = bbtags[bbnumber] + theSelection + bbtags[bbnumber+1]
				}
			}
			// Comprobar que todos los campo están
			for (n=1;n<=totalCampos;n++) {
				campo = form1['texto'+n]
				if (campo != '[object]') {
					alert('Por favor, seleccione sólamente dentro de los campos de texto.\nLas modificaciones se han cancelado.')
					//this.refresh()
					location.reload()
					unError = true
				}
			}
		} else {
//			alert("Ha seleccionado una herramienta de: '"+nombreCampo+"'\nLa selección está declara en el campo: '"+focoActual+"'")
			// focoActual pasa ha ser el del que ha tocado la herramienta. (se manda el foco a el)
			focoActual = nombreCampo
		}
		if(!unError){
			document[nombreForm][nombreCampo].focus();
		}


		theSelection = '';
		return;
	}

	// Find last occurance of an open tag the same as the one just clicked
	for (i = 0; i < bbcode.length; i++) {
		if (bbcode[i] == bbnumber+1) {
			bblast = i;
			donotinsert = true;
		}
	}
	if (donotinsert) {		// Close all open tags up to the one just clicked & default button names
		while (bbcode[bblast]) {
				butnumber = arraypop(bbcode) - 1;
				document[nombreForm][nombreCampo].value += bbtags[butnumber + 1];
				buttext = eval('document[nombreForm].addbbcode' + butnumber + '.value');
				eval('document[nombreForm].addbbcode' + butnumber + '.value ="' + buttext.substr(0,(buttext.length - 1)) + '"');
				imageTag = false;
			}
			document[nombreForm][nombreCampo].focus();
			return;
	} else { // Open tags
		alert("Por favor, seleccione el texto al que desea aplicar el formato.")
/*
		if (imageTag && (bbnumber != 14)) {		// Close image tag before adding another
			document[nombreForm][nombreCampo].value += bbtags[15];
			lastValue = arraypop(bbcode) - 1;	// Remove the close image tag from the list
			document[nombreForm].addbbcode14.value = "Img";	// Return button back to normal state
			imageTag = false;
		}

		// Open tag
		document[nombreForm][nombreCampo].value += bbtags[bbnumber];
		if ((bbnumber == 14) && (imageTag == false)) imageTag = 1; // Check to stop additional tags after an unclosed image tag
		arraypush(bbcode,bbnumber+1);
		eval('document[nombreForm].addbbcode'+bbnumber+'.value += "*"');
		document[nombreForm][nombreCampo].focus();
		return;
	*/
	}
	storeCaret(document[nombreForm][nombreCampo]);

}

// Insert at Claret position. Code from
// http://www.faqts.com
function storeCaret(textEl) {
	if (textEl.createTextRange) textEl.caretPos = document.selection.createRange().duplicate();
}
//-->
</script> 
<%
Sub pinta_bb ( texto )
'	texto = Replace (texto,"<BR>","<br>")
'	texto = Replace (texto,"<Br>","<br>")
'	texto = Replace (texto,"<bR>","<br>")
'	texto = Replace (texto,"<br>", vbCrLf)
	Response.write (texto)
End Sub


function formImagen (n,nombreCampo,nombreFoto,ruta,anchoMax,altoMax,anchoMin,altoMin,margen,novisible,pie,enlace,enlaceventana)

	if novisible="" then
		novisible=0
	end if
%>
<form action="cambiarFoto.asp" method="post" enctype="multipart/form-data" name="form<%=n%>">
<input type="hidden" name="rutavuelta" value="<%=ruta%>">
<input type="hidden" name="nodo" value="imagen<%=n%>">
<input type="hidden" name="anchoMax" value=<%=anchoMax%>>
<input type="hidden" name="altoMax" value=<%=altoMax%>>
<input type="hidden" name="anchoMin" value=<%=anchoMin%>>
<input type="hidden" name="altoMin" value=<%=altoMin%>>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="0">
	<tr>
	<td align="left" valign="top">
	<%
	if nombreFoto <> "" then
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		laRutaImagen = ruta &"/"& nombreFoto
		if fso.FileExists(server.MapPath("../..")&"/"&session("idioma")&laRutaImagen) then%>
			<table  border="1" cellpadding="0" cellspacing="10" bordercolor="#ECE9D8">
              <tr>
                <td align="center" valign="middle"><img src="../../<%=session("idioma")%>/<%=laRutaImagen%>" alt="<%=nombreFoto%>" border="0"></td>
              </tr>
            </table>
			<br>
		<%else%>
	        <br><b>Error</b>: La imagen indicada no existe. Introduzca una nueva.
		<%end if
	else%>
		<br>Introduzca <nobr>una imagen.</nobr><br>
		<br>
	<%end if%></td>
	<td width="100%" align="center" valign="middle">
	
	  <table  border="0" cellpadding="1" cellspacing="0" bgcolor="#990000">
      <tr>
        <td><table width="100%" border="0" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
          <%if anchoMax > 0 then%>
          <tr>
            <td align="right">Ancho m&aacute;ximo:</td>
            <td><b><%=anchoMax%> px</b></td>
          </tr>
          <%end if
	if altoMax > 0 then%>
          <tr>
            <td align="right">Alto m&aacute;ximo:</td>
            <td><b><%=altoMax%> px </b></td>
          </tr>
          <%end if
	if anchoMin > 0 then%>
          <tr>
            <td align="right">Ancho m&iacute;nimo:</td>
            <td><b><%=anchoMin%> px</b></td>
          </tr>
          <%end if
	if altoMin > 0 then%>
          <tr>
            <td align="right">Alto m&iacute;nimo:</td>
            <td><b><%=altoMin%> px</b></td>
          </tr>
          <%end if%>
        </table></td>
      </tr>
    </table>	</td>
	</tr>
</table>


<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td>
	<fieldset>
	<legend><%=elComentario%></legend>
	<input type="hidden" name="comentario_imagen" value="<%=elComentario%>">
	<table width="100%" border="0" align="center" cellpadding="8" cellspacing="0" style="border=1px">
		<%if elPie <> "-1" then%>
			<tr>
			<td colspan="2">
			<table width="100%"  border="0" cellspacing="0" cellpadding="1">
				<tr>
				<td>Enlace:</td>
				<td>&nbsp;</td>
				</tr>
				<tr>
				<td width="100%"><input name="enlace" type="text" id="enlace" style="width:100%" value="<%=enlace%>" maxlength="250"></td>
				<td><select name="enlaceventana" id="enlaceventana">
				<option value="_self" <%if enlaceventana = "_self" then Response.Write "selected" end if%>>Misma ventana</option>
				<option value="_blank" <%if enlaceventana = "_blank" then Response.Write "selected" end if%>>Ventana nueva</option>
				</select></td>
				</tr>
			</table></td>
			</tr>
			<tr>
			<td colspan="2">Pie de foto:<br><textarea name="pie" cols="50" rows="3" wrap="virtual" style="width:100%"><%=elPie%></textarea></td>
			</tr>
		<%else%>
			<input type="hidden" name="pie" value="-1">
		<%end if%>
		<tr>
		<td colspan="2" valign="top"><input type="hidden" name="fotoAnterior" value="<%=nombreFoto%>">
		Escoja su nueva imagen:<br>
		<input name="fotoNueva" type="file" style="height=20px;width:100%" size="45"></td>
		</tr>
		<%if novisible<> "-1" or margen <> "-1" then%>
		<tr>
		<td>
		<%if novisible <> "-1" then%>
			<input name="novisible" type="checkbox" id="novisible" value="1" <%if novisible="1" then response.write("checked") end if%>>No visible
		<%else%>
			<input type="hidden" name="novisible" value="-1">
		<%end if
		if margen <> "-1" then%>
			<input name="margen" type="checkbox" id="margen" value="1" <%if margen="1" then response.write("checked") end if%>>Con borde
		<%else%>
			<input type="hidden" name="margen" value="-1">
		<%end if%></td>
		<td align="right"><input type="submit" style="height=20px" value="Enviar" class="botonAdmin"></td>
		</tr>
	<%end if%>
</table>
	  
	  </fieldset>
	  
	  </td>
    </tr>
  </table>
</form>
<%
	Set fso=Nothing
	end function

function textAreaBbcode (nombreForm,nombreCampo,contenido,ancho,alto,legend)
	'para que el campo no sea inmenso, con 30 tiene de sobra
	if alto>=30 then
		alto=30
	end if
	%>
	<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
	
<fieldset>
	<legend><%=legend%></legend>

	<table width="100%"  border="0" align="center" cellpadding="6" cellspacing="0">
		<tr>
		<td align="center"><textarea name="<%=nombreCampo%>" cols="" rows="<%=alto%>" wrap="virtual" class="campoAdmin" style="width:100%" onFocus="focoActual = this.name" onSelect="storeCaret(this);" onClick="storeCaret(this);" onKeyUp="storeCaret(this);"><%pinta_bb contenido%></textarea></td>
		</tr>
	</table>
	</fieldset>
	
	</td>
  </tr>
</table>

	
	<br>
	<script language="javascript1.2">
	editor_generate('<%=nombreCampo%>');
	</script>
<%end function%>