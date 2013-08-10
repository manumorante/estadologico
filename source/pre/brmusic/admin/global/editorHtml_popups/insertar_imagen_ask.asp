<!--#include virtual="/datos/inc_config_gen.asp" -->
<%
conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("/"& c_s &"datos/esp/fotospagina/fotospagina.mdb")
%>


<html id=dlgImage STYLE="width: 550px; height: 225px; ">
<head>
<title>Insertar imagen de biblioteca</title>
<SCRIPT defer>

function _CloseOnEsc() {
  if (event.keyCode == 27) { window.close(); return; }
}

function _getTextRange(elm) {
  var r = elm.parentTextEdit.createTextRange();
  r.moveToElementText(elm);
  return r;
}

window.onerror = HandleError

function HandleError(message, url, line) {
  var str = "An error has occurred in this dialog." + "\n\n"
  + "Error: " + line + "\n" + message;
  alert(str);
  window.close();
  return true;
}

function Init() {
  var elmSelectedImage;
  var htmlSelectionControl = "Control";
  var globalDoc = window.dialogArguments;
  var grngMaster = globalDoc.selection.createRange();
  
  // event handlers  
  document.body.onkeypress = _CloseOnEsc;
  btnOK.onclick = new Function("btnOKClick()");

  txtFileName.fImageLoaded = false;
  txtFileName.intImageWidth = 0;
  txtFileName.intImageHeight = 0;

  if (globalDoc.selection.type == htmlSelectionControl) {
    if (grngMaster.length == 1) {
      elmSelectedImage = grngMaster.item(0);
      if (elmSelectedImage.tagName == "IMG") {
        txtFileName.fImageLoaded = true;
        if (elmSelectedImage.src) {
          txtFileName.value          = elmSelectedImage.src.replace(/^[^*]*(\*\*\*)/, "$1");  // fix placeholder src values that editor converted to abs paths
          txtFileName.intImageHeight = elmSelectedImage.height;
          txtFileName.intImageWidth  = elmSelectedImage.width;
          txtVertical.value          = elmSelectedImage.vspace;
          txtHorizontal.value        = elmSelectedImage.hspace;
          txtBorder.value            = elmSelectedImage.border;
          txtAltText.value           = elmSelectedImage.alt;
          selAlignment.value         = elmSelectedImage.align;
        }
      }
    }
  }
  txtFileName.value = txtFileName.value || "http://";
  txtFileName.focus();
}

function _isValidNumber(txtBox) {
  var val = parseInt(txtBox);
  if (isNaN(val) || val < 0 || val > 999) { return false; }
  return true;
}

function btnOKClick() {
  var elmImage;
  var intAlignment;
  var htmlSelectionControl = "Control";
  var globalDoc = window.dialogArguments;
  var grngMaster = globalDoc.selection.createRange();
  
  // error checking

  if (!txtFileName.value || txtFileName.value == "http://") { 
    alert("Image URL must be specified.");
    txtFileName.focus();
    return;
  }
  if (txtHorizontal.value && !_isValidNumber(txtHorizontal.value)) {
    alert("Horizontal spacing must be a number between 0 and 999.");
    txtHorizontal.focus();
    return;
  }
  if (txtBorder.value && !_isValidNumber(txtBorder.value)) {
    alert("Border thickness must be a number between 0 and 999.");
    txtBorder.focus();
    return;
  }
  if (txtVertical.value && !_isValidNumber(txtVertical.value)) {
    alert("Vertical spacing must be a number between 0 and 999.");
    txtVertical.focus();
    return;
  }

  // delete selected content and replace with image
  if (globalDoc.selection.type == htmlSelectionControl && !txtFileName.fImageLoaded) {
    grngMaster.execCommand('Delete');
    grngMaster = globalDoc.selection.createRange();
  }
    
  idstr = "\" id=\"556e697175657e537472696e67";     // new image creation ID
  if (!txtFileName.fImageLoaded) {
    grngMaster.execCommand("InsertImage", false, idstr);
    elmImage = globalDoc.all['556e697175657e537472696e67'];
    elmImage.removeAttribute("id");
    elmImage.removeAttribute("src");
    grngMaster.moveStart("character", -1);
  } else {
    elmImage = grngMaster.item(0);
    if (elmImage.src != txtFileName.value) {
      grngMaster.execCommand('Delete');
      grngMaster = globalDoc.selection.createRange();
      grngMaster.execCommand("InsertImage", false, idstr);
      elmImage = globalDoc.all['556e697175657e537472696e67'];
      elmImage.removeAttribute("id");
      elmImage.removeAttribute("src");
      grngMaster.moveStart("character", -1);
      txtFileName.fImageLoaded = false;
    }
    grngMaster = _getTextRange(elmImage);
  }

  if (txtFileName.fImageLoaded) {
    elmImage.style.width = txtFileName.intImageWidth;
    elmImage.style.height = txtFileName.intImageHeight;
  }

  if (txtFileName.value.length > 2040) {
    txtFileName.value = txtFileName.value.substring(0,2040);
  }
  
  elmImage.src = txtFileName.value;
  
  if (txtHorizontal.value != "") { elmImage.hspace = parseInt(txtHorizontal.value); }
  else                           { elmImage.hspace = 0; }

  if (txtVertical.value != "") { elmImage.vspace = parseInt(txtVertical.value); }
  else                         { elmImage.vspace = 0; }
  
  elmImage.alt = txtAltText.value;

  if (txtBorder.value != "") { elmImage.border = parseInt(txtBorder.value); }
  else                       { elmImage.border = 0; }

  elmImage.align = selAlignment.value;
  grngMaster.collapse(false);
  grngMaster.select();
  window.close();
}

//
			function lanzapo_win(theURL,winName,ancho,alto,barras) { 
				var winl = (screen.width - ancho) / 2;
				var wint = (screen.height - alto) / 2;
				var paramet='top='+wint+',left='+winl+',width='+ancho+',height='+alto+',scrollbars='+barras+'';
				var splashWin=window.open(theURL,winName,paramet);
				splashWin.focus();
			}
//
function verLista(){
	showModalDialog("../../visores/visor_imagenes.asp", "", "resizable: no; help: no; status: no; scroll: no; ");
}

//
function ampliarFoto(){
	var ancho = 300
	var alto = 250
	var barras = 0
	var list = document.getElementById("lista")
	var winl = (screen.width - ancho) / 2;
	var wint = (screen.height - alto) / 2;
	var paramet='top='+wint+',left='+winl+',width='+ancho+',height='+alto+',scrollbars='+barras+'';
	var nw = window.open("","VistaPrevia",paramet)
	nw.document.write('<html>\n<head>\n<title>Vista previa de imagen</title>\n<scr'+'ipt>function tamano(w,h) {\nif (w < screen.width-100 && h < screen.height-100) {\nvar winl = (screen.width - w) / 2;\nvar wint = (screen.height - h) / 2;\nmoveTo(winl,wint);\nresizeTo(w+30,h+60);\n} else {\nw = screen.width - 100;\nh = screen.height - 150;\nvar winl = (screen.width - w) / 2;\nvar wint = (screen.height - h) / 2;\nmoveTo(winl,wint);\nresizeTo(w+30,h+60);\nalert("La imagen es mas grande que la pantalla y se ve cortada.");\n}\n}\n</scr'+'ipt>\n</head>\n<body>\n\n<img src="../../../datos/esp/fotospagina/fotos/'+ list.value +'" onLoad="tamano(this.width,this.height)">\n\n</body>\n</html>')
}

//
function changeLista(c) {
	var dire = document.getElementById("txtFileName")
	dire.value = "/<%=c_s%>datos/esp/fotospagina/fotos/"+c
}
</SCRIPT>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><style type="text/css">
<!--
body {
	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
	background-color: #ECE9D8;
}
body,td,th {
	font-size: 8pt;
	font-family: MS Shell Dlg;
}
-->
</style></head>
<body class="bodyAdmin"  id="bdy" onLoad="Init()">
<table width="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><fieldset>
    <legend>Insertar imagen de biblioteca</legend>
    <table width="100%" border="0" cellpadding="4" cellspacing="0">
      <tr>
        <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td>Escojer:&nbsp;</td>
              <td>
			  <%
	sql = "SELECT R_FOTO, R_TITULO FROM REGISTROS WHERE R_FOTO <> ''"
	set re = Server.CreateObject("ADODB.Recordset")
	re.ActiveConnection = conn_
	re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3 : re.Open()
			  %>
	<select name="lista" id="lista" onChange="changeLista(this.value)">
	<option value="">Lista ...</option>
		<%
		n=0
		while not re.eof
		n= n+1%>
		<option value="<%=re("R_FOTO")%>"><%=n%>) <%=re("R_TITULO")%></option>
		<%
			re.movenext
		wend%>
	</select>
	<input name="" type="button" onClick="ampliarFoto()" value="Ver">
<%
	re.close()
	set re = nothing
%></td>
            </tr>
            <tr>
              <td>Url:&nbsp;</td>
              <td width="100%"><input name="txtFileName" type="text" id="txtFileName" style="width:100%" tabindex="10" onFocus="select()"></td>
            </tr>
            <tr>
              <td>Alt:&nbsp;</td>
              <td><input name="txtAltText" type="text" id="txtAltText" style="width:100%" tabindex="15" onFocus="select()"></td>
            </tr>
        </table></td>
      </tr>
    </table>
    </fieldset>
      <table width="100%"  border="0" cellpadding="1" cellspacing="0">
        <tr>
          <td><fieldset>
          <legend>Disposición</legend>
          <table width="100%" border="0" cellpadding="4" cellspacing="0">
            <tr>
              <td><table border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>Alinear:&nbsp;</td>
                    <td>
                      <SELECT size="1" ID="selAlignment" tabIndex="20">
                        <OPTION id=optNotSet value=""> Not set </OPTION>
                        <OPTION id=optLeft value=left> Left </OPTION>
                        <OPTION id=optRight value=right> Right </OPTION>
                        <OPTION id=optTexttop value=textTop> Texttop </OPTION>
                        <OPTION id=optAbsMiddle value=absMiddle> Absmiddle </OPTION>
                        <OPTION id=optBaseline value=baseline SELECTED> Baseline </OPTION>
                        <OPTION id=optAbsBottom value=absBottom> Absbottom </OPTION>
                        <OPTION id=optBottom value=bottom> Bottom </OPTION>
                        <OPTION id=optMiddle value=middle> Middle </OPTION>
                        <OPTION id=optTop value=top> Top </OPTION>
                    </SELECT></td>
                  </tr>
                  <tr>
                    <td>Borde:&nbsp;</td>
                    <td><input name="txtBorder" type="text" id="txtBorder" tabindex="15" onFocus="select()" size="3" maxlength="3"></td>
                  </tr>
              </table></td>
            </tr>
          </table>
          </fieldset></td>
          <td><fieldset>
          <legend>Espaciado</legend>
          <table width="100%" border="0" cellpadding="4" cellspacing="0">
            <tr>
              <td><table border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td>Horizontal:&nbsp;</td>
                    <td>                      <input name="txtHorizontal" type="text" size="3" maxlength="3"></td>
                  </tr>
                  <tr>
                    <td>Vertical:&nbsp;</td>
                    <td><input name="txtVertical" type="text" id="txtVertical" tabindex="15" onFocus="select()" size="3" maxlength="3"></td>
                  </tr>
              </table></td>
            </tr>
          </table>
          </fieldset></td>
        </tr>
      </table>    </td>
    <td align="right" valign="top"><table  border="0" cellspacing="0" cellpadding="2">
      <tr>
        <td><input name="btnOK" type="submit" id="btnOK" style="width:7em" value="OK"></td>
      </tr>
      <tr>
        <td><input name="btnCancel" type="reset" id="btnCancel" style="width:7em" onClick="window.close();" value="Cancelar"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</body>
</html>