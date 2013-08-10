<html><!-- InstanceBegin template="/Templates/aneg.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Asociación Nacional de Especialistas y Expertos en Gerontagogía</title>
<!-- InstanceEndEditable -->
<%
	randomize
	ale = Int(Rnd * (totalRegistros))+1
%>
<link href="/arch/estilos.css?ale=<%=ale%>" rel="stylesheet" type="text/css" media="screen">
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<script language="javascript" type="text/javascript" src="/arch/js.js"></script>
</head>
<body bgcolor="#7BC548" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include virtual="/inc/config.asp" -->
<div align="center">
  <!-- InstanceBeginEditable name="Inicio" -->
  <!--#include virtual="/admin/inc_rutinas.asp" -->
<!-- InstanceEndEditable -->
  <!--#include virtual="/inc/cabecera.asp" -->
  <table width="779" border="0" cellspacing="0" cellpadding="0">
  <tr>

    <td align="left" valign="top" bgcolor="#FFFFFF"><!-- InstanceBeginEditable name="Medio" -->
<!--#include virtual="/Asociados/subsecciones.asp" -->

<!-- InstanceEndEditable -->
      <table width="100%" border="0" cellspacing="0" cellpadding="30">
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF" class="bg_centro"><!-- InstanceBeginEditable name="Cuerpo" --><span class="titulo">Noticias</span><br>
              <br>

              <%cualid="noticias"%>
			  <!--#include file="visores/inc_visor_conciertos.asp" -->
          <!-- InstanceEndEditable --></td></tr>
      </table></td></tr>
</table>
  <!--#include virtual="/inc/pie.asp" -->
<!-- InstanceBeginEditable name="Fin" --><!-- InstanceEndEditable --></div>
</body>
<!-- InstanceEnd --></html>
