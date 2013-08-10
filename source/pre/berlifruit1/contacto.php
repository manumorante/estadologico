<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/index.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Description" content="Berli Fruit - Grupo Agroalimentario, Frutas y verduras, Quesos, Productos ibéricos, Vinos" />
<meta name="Keywords" content="berli fruit, grupo agroalimentario, frutas y verduras, quesos, productos ibéricos, vinos" />
<meta name="title" content="Berli Fruit - Grupo Agroalimentario" />
<meta name="DC.Language" scheme="RFC1766" content="Spanish" />
<meta name="Revisit-after" content="15 days" />
<meta name="robots" content="ALL,FOLLOW" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Berli Fruit - Grupo Agroalimentario</title>
<!-- InstanceEndEditable -->
<link href="css/comun.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
</head>
<body bgcolor="#FFFFFF">
<div id="Cuerpo">
  <div id="Cabecera">
    <div id="Logo"><a href="index.html"><img src="arch/logo.gif" alt="Berli Fruit" width="79" height="81" border="0" longdesc="Grupo Agroalimentario" /></a></div>
    <div id="MenuSuperior"><a href="#">Berli Fruit, Grupo Agroalimentario</a> | <a href="contacto.html">Contacto</a></div>
    <div id="Menu">
	<ul>
	<li><a href="productos.html">Productos</a></li>
	<li><a href="la_empresa.html">La empresa</a></li>
	<li><a href="localizacion.html">Localizaci&oacute;n</a></li>
	<li><a href="contacto.html">Contacto</a></li>
	</ul>
    </div>
  </div>
  <!-- InstanceBeginEditable name="Cabecera" -->
  <div id="Central2"></div>
  <!-- InstanceEndEditable -->
  <div id="CajaASeccion">
    <div id="CajaBSeccion">
      <h1><!-- InstanceBeginEditable name="Título" -->Contacte con nosotros<!-- InstanceEndEditable --></h1>
	  
	  
      <!-- InstanceBeginEditable name="Cuerpo" -->
      <div id="ColumnaIzq">
        <p>Contacte de forma f&aacute;cil y r&aacute;pida tan solo<br />
        con rellenar el siguiente formulario.</p>
         <? 
if (!$HTTP_POST_VARS){ 
?>
         <form name="enviar" action="contacto.php" method="post">
		<table cellpadding="0" cellspacing="0" id="TablaContacto">
          <tr>
            <th>Su nombre: </th>
          </tr>
          <tr>
            <td><input name="nombre" type="text" class="campo" id="nombre" maxlength="50" /></td>
          </tr>
          <tr>
            <th>Su empresa: </th>
          </tr>
          <tr>
            <td><input name="empresa" type="text" class="campo" id="empresa" maxlength="50" /></td>
          </tr>
          <tr>
            <th>Direcci&oacute;n: </th>
          </tr>
          <tr>
            <th><input name="direccion" type="text" class="campo" id="direccion" maxlength="50" /></th>
          </tr>
          <tr>
            <th>Tel&eacute;fono:</th>
          </tr>
          <tr>
            <td><input name="telefono" type="text" class="campo" id="telefono" maxlength="50" /></td>
          </tr>
          <tr>
            <th>E-mail:</th>
          </tr>
          <tr>
            <td><input name="mail" type="text" class="campo" id="mail" maxlength="50" /></td>
          </tr>
          <tr>
            <th>Mensaje:</th>
          </tr>
          <tr>
            <td><textarea name="mensaje" cols="" rows="5" wrap="virtual" class="campo" id="mensaje"></textarea></td>
          </tr>
          <tr align="right">
            <td height="45" valign="bottom"><input type=submit name="Submit" value="Enviar" />			</td>
          </tr>
        </table>
		<? 
}else{ 
    //Estoy recibiendo el formulario, compongo el cuerpo 
    $cuerpo = "Formulario enviado\n"; 
    $cuerpo .= "Nombre: " . $HTTP_POST_VARS["nombre"] . "\n"; 
    $cuerpo .= "Empresa: " . $HTTP_POST_VARS["empresa"] . "\n"; 
	 $cuerpo .= "Dirección: " . $HTTP_POST_VARS["direccion"] . "\n"; 
    $cuerpo .= "Telefono: " . $HTTP_POST_VARS["telefono"] . "\n"; 
    $cuerpo .= "Email: " . $HTTP_POST_VARS["mail"] . "\n"; 
    $cuerpo .= "Comentario: " . $HTTP_POST_VARS["mensaje"] . "\n"; 
	
	$cabeceras .= "From: e-mail enviado desde el portal www.berlifruit.com por el Sr./a:<'$nombre'> de la empresa<'$empresa>\r\n"; 
	$cabecerausuario .= "From: www.berlifruit.com - berlifruit@berlifruit.com\r\n";
//mando el correo... 


    //mando el correo... 
    mail("berlifruit@berlifruit.com","Formulario recibido",$cuerpo,$cabeceras); 
    echo "<br><br><br><br>"; 
$respuesta ='Formulario recibido, pronto nos pondremos en contacto con usted. Berlifruit.' ; 
mail("$mail","Formulario recibido",$cuerpo ."\n". $respuesta,$cabecerausuario); 
     echo "<div align='center'>Mensaje enviado correctamente.</div>";
} 
?>
		</form>
      </div>
	<div id="ColumnaDer">
	  <p><img src="arch/fruta.jpg" alt="Fruta" width="319" height="136" /></p>
	  <p><strong>Tel&eacute;fonos</strong>:<br />
	    95 513 40 02<br />
	    95 513 40 03</p>
	  <p><strong>Fax</strong>: <br />
	    95 513 40 01</p>
	  <p><strong>Direcci&oacute;n</strong>:<br />
	    Ctra. Madrid-C&aacute;diz Km. 445<br />
	    41400 &Eacute;CIJA (Sevilla)<br />
	    <a href="mailto:berlifruit@berlifruit.com"><br />
	    </a><strong>E-mail:</strong>	    <a href="mailto:berlifruit@berlifruit.com">berlifruit@berlifruit.com</a><br />
	  </p>
	  </div>
	  <!-- InstanceEndEditable -->
	  
	  
	</div>
	<div class="Borrar"></div>
  </div>
  <div id="Pie"> Copyright &copy; 2006 <strong>Berli Fruit</strong>, Grupo Agroalimentario | Tel&eacute;fono: 95 513 40 02 - 95 513 40 03 - Fax: 95 513 40 01</div>
  <div id="Pie2"></div>
</div>
</body>
<!-- InstanceEnd --></html>
