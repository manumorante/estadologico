<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>No acceso</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../estilos.css" rel="stylesheet" type="text/css">
<script>
	// Si esta página está continida en frames se saldrá.
	if (parent.frames.length >0){
		top.location.href="noacceso.asp"
	}

	// Ir a la pagina inicial de aSkipper (puesto que está logeado).
	function paginaInicial() {
		top.location.href='../contenedor.htm'
	}
</script>
</head>
<body leftmargin="0" topmargin="0" class="bodyAdmin">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center" valign="middle"><table border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#990000">
        <tr>
          <td><table border="0" align="center" cellpadding="8" cellspacing="0" bgcolor="#fafafa">
              <tr>
                <td bgcolor="#990000"><font color="#FFFFFF" size="2"><b>Usted
                      no tiene permiso para acceder a esta zona</b></font></td>
              </tr>
              <tr>
                <td>Si cree que esto es un error p&oacute;ngase en contacto con
                  nosotros en la secci&oacute;n <a href="../../esp/index.asp?secc=/contacto">Contacto</a> de
                  la p&aacute;gina principal o vuelva a la p&aacute;gina inicial de
                  administraci&oacute;n para validarse pulsando en el siguiente enlace. </td>
              </tr>
              <tr>
                <td align="center"><a href="../">P&aacute;gina inicial</a>                 </td>
              </tr>
            </table>
          </td>
        </tr>
      </table></td>
    </tr>
</table>
  <br>
  <br>
</div>
</body>
</html>
