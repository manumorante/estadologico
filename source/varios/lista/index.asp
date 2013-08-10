<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Index</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body.table,tr,td {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9pt;
	color: #333333;
}
a:link {
	color: #666666;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	color: #000000;
	text-decoration: underline;
}
a:active {
	color: #000000;
	text-decoration: none;
}
-->
</style>
</head>

<body bgcolor="#ECE9D8" vlink="#666666">
<%
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set dir = fso.getFolder(server.MapPath("."))
	for each a in dir.files
		Response.Write "<br>" & a
	next
	%>
	

	<table  border="0" align="center" cellpadding="1" cellspacing="0" bordercolor="#666666" bgcolor="#ECE9D8">
      <tr>
        <td align="center">
		<fieldset>
	<legend>Lista de carpetas</legend>
	<table  border="0" cellpadding="10" cellspacing="0" bgcolor="#F8F7F1">
          <tr>
            <td><table border="0" cellpadding="1" cellspacing="0">
                <tr>
                  <td valign="middle"><a href="w:/" target="_blank"><img src="../arch/ico_disco_duro.gif" width="16" height="11" border="0"></a></td>
                  <td><a href="w:/" target="_blank">Webs</a></td>
                </tr>
                <%for each a in dir.subFolders
		if inStr(a.name,"vti") <=0 and inStr(a.name,"private") <=0 and a.name <> "RECYCLER" and a.name <> "System Volume Information" then
			alt=ucase(a.name)
			lineas = 0
			for each b in a.subFolders
				if lineas < 30 and inStr(b.name,"vti") <=0 and inStr(b.name,"private") <=0 and b.name <> "RECYCLER" and b.name <> "System Volume Information" then
					alt = alt & vbCrlf & "  "&b.name
					lineas = lineas + 1
				end if
			next
			for each b in a.files
				if lineas <30 and inStr(b.name,"vti") <=0 and inStr(b.name,"._") <=0 and inStr(b.name,".DS") <=0 then
					alt = alt & vbCrlf & "  "&b.name
					lineas = lineas + 1
				end if
			next
			if lineas >=30 then
				alt = alt & vbCrlf & " ... "
			end if
		%>
                <tr>
                  <td valign="middle"><a href="w:/<%=a.name%>" target="_blank"><img src="../arch/ico_carpeta.gif" alt="<%=alt%>" width="18" height="15" border="0"></a></td>
                  <td><a href="<%=a.name%>" target="_blank"><%=a.name%></a></td>
                </tr>
                <%end if
	next%>
            </table></td>
          </tr>
        </table>
        </fieldset>		</td>
      </tr>
</table>
</body>
</html>
