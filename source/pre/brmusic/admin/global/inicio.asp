<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%dim msgerror, unerror%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/inc_rutinas.asp" -->
<!--#include file="inc_rutinas.asp" -->
<%cualid = session("cualid")%>
<!--#include file="inc_inicia_xml.asp" -->
<%inicia_xml%>
<!--#include file="inc_conn.asp" -->
<html>
<head>
<title>Administración</title>
<script src="../rutinas.js" language="javascript" type="text/javascript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	background-color: #f5f5f5;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body class="bodyAdmin">
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><br>
<%if request.QueryString("msg") <> "" then%>
<table width="370" border="0" cellspacing="0" cellpadding="1">
  <tr>
    <td><%=request.QueryString("msg")%></td>
  </tr>
</table>
<%end if%>

<%if request.QueryString("msgerror") <> "" then%>
<table width="370" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td bgcolor="#FF0000"><b><font color="#FFFFFF">ERROR: <br>
      <%=request.QueryString("msgerror")%></font></b></td>
  </tr>
</table>
<%end if%>

<%if not unerror then%>
	<font color="849ace"><%genXml()%></font>
<%end if%>

<%select case request.QueryString("ac")%>
<%case "editarpermisos"
	id = numero(request.QueryString("id"))
	if id >0 then%>
		<script language="javascript" type="text/javascript">
		<!--
			winPop("usuarios_personalizados.asp?ac=editar&id=<%=id%>","Permisos",475,575,1)
		//-->
		</script>
	<%end if
end select%>


<%if unerror then
	Response.Write "<b>Error</b>: " & msgerror 
end if%>
</td>
  </tr>
</table>
</body>
</html>
