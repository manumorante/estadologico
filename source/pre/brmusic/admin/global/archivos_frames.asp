<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Skipper</title>
</head>

<%
	if not config_iexplorer then
		rows = "15%,85%"
	else
		rows = "100%,0%"
	end if

%>

<frameset rows="<%=rows%>" frameborder="NO" border="0" framespacing="0">
  <frame src="archivos.asp?<%=request.QueryString()%>" name="archivos">
  <frame src="archivos_form.asp?<%=request.QueryString()%>" name="archivos_form">
</frameset>
<noframes><body>
</body></noframes>
</html>
