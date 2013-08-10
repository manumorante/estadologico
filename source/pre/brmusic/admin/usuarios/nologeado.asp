<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>No acceso</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../estilos.css" rel="stylesheet" type="text/css">
<script>
	// Si esta página está continida en frames se saldrá.
	if (parent.frames.length >0){
		top.location.href="nologeado.asp"
	}
	
	// Ir a la página de validación (puesto que no está validado).
	function entrar() {
		top.location.href='validar.asp'
	}
</script>
</head>
<body class="bodyAdmin">

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><table border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#990000">
        <tr>
          <td><table border="0" align="center" cellpadding="8" cellspacing="0" bgcolor="#fafafa">
              <tr>
                <td bgcolor="#990000"><font color="#FFFFFF" size="2"><b>Usted
                      no est&aacute; validado en el sistema</b></font></td>
              </tr>
              <tr>
                <td>Para validarse pulse sobre el siguiente bot&oacute;n e introduzca
                  su nombre de usuario y contrase&ntilde;a. <b></b><br>
                  </td>
              </tr>
              <tr>
                <td align="center"><input type="button" onClick="entrar()" value="Ir a validaci&oacute;n">
                  </td>
              </tr>
              </table>
          </td>
        </tr>
        </table>
    </td>
  </tr>
    </table>
</body>
</html>
